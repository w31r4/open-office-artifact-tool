#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { generateOfficeInput } from "./agent-eval-office-fixtures.mjs";
import { gradeOfficeCase } from "./agent-eval-office-graders.mjs";
import { gradePdfCase } from "./agent-eval-pdf-graders.mjs";
import { checkReferenceSkillSync, REFERENCE_SKILLS_ROOT } from "./reference-skill-sync.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
// PromptBench remains PDF-led while it grows independently graded workflows
// for the four artifact families. The durable product invariant is an absolute
// PDF majority, not a fixed percentage that would force unrelated filler
// prompts whenever a meaningful Office vertical slice is added.
export const MINIMUM_PDF_CASE_SHARE = 0.5;
const casesPath = path.join(repoRoot, "evals", "cases.jsonl");
const families = new Set(["pdf", "documents", "spreadsheets", "presentations"]);
const skillsByFamily = new Map([
  ["pdf", "pdf"],
  ["documents", "documents"],
  ["spreadsheets", "spreadsheets"],
  ["presentations", "presentations"],
]);
const savePolicies = new Set(["read-only", "none", "rewrite", "incremental", "sanitize"]);
const mimeMagic = new Map([
  ["application/pdf", Buffer.from("%PDF-")],
  ["application/vnd.openxmlformats-officedocument.wordprocessingml.document", Buffer.from("PK")],
  ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", Buffer.from("PK")],
  ["application/vnd.openxmlformats-officedocument.presentationml.presentation", Buffer.from("PK")],
]);

function fail(message) {
  throw new Error(message);
}

function parseArgs(argv) {
  const [command = "help", ...rest] = argv;
  const flags = new Map();
  const positionals = [];
  for (let index = 0; index < rest.length; index += 1) {
    const value = rest[index];
    if (!value.startsWith("--")) {
      positionals.push(value);
      continue;
    }
    const [key, inline] = value.slice(2).split("=", 2);
    if (inline !== undefined) flags.set(key, inline);
    else if (rest[index + 1] && !rest[index + 1].startsWith("--")) flags.set(key, rest[++index]);
    else flags.set(key, true);
  }
  return { command, positionals, flags };
}

async function loadRecords() {
  const source = await fs.readFile(casesPath, "utf8");
  return source.split(/\r?\n/).map((line) => line.trim()).filter(Boolean).map((line, index) => {
    try {
      return JSON.parse(line);
    } catch (error) {
      fail(`evals/cases.jsonl:${index + 1}: ${error.message}`);
    }
  });
}

async function loadSuite() {
  const records = await loadRecords();
  const suiteRecords = records.filter((record) => record.record === "suite");
  if (suiteRecords.length !== 1) fail(`expected exactly one suite record, found ${suiteRecords.length}`);
  return { suite: suiteRecords[0], cases: records.filter((record) => record.record === "case") };
}

function safeRelative(target, label) {
  if (typeof target !== "string" || !target || target.includes("\0") || path.isAbsolute(target)) fail(`${label} must be a non-empty relative path`);
  const normalized = path.normalize(target);
  if (normalized === "." || normalized === ".." || normalized.startsWith(`..${path.sep}`)) fail(`${label} escapes the workspace: ${target}`);
  return normalized;
}

function isUnder(normalized, directory) {
  return normalized.startsWith(`${directory}${path.sep}`) && normalized.length > directory.length + 1;
}

function validateSuite(suite, cases) {
  const errors = [];
  const ids = new Set();
  const statuses = new Set(suite.statuses || []);
  const outcomes = new Set(suite.outcomes || []);
  const weights = Object.values(suite.weights || {}).reduce((sum, value) => sum + value, 0);
  if (weights !== 100) errors.push(`suite weights must total 100, found ${weights}`);
  if (!Number.isInteger(suite.defaultTrials) || suite.defaultTrials < 1) errors.push("defaultTrials must be a positive integer");
  for (const item of cases) {
    const prefix = item.id || "<missing-id>";
    const inputPaths = new Set();
    const deliverablePaths = new Set();
    if (!/^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(item.id || "")) errors.push(`${prefix}: invalid id`);
    if (ids.has(item.id)) errors.push(`${prefix}: duplicate id`);
    ids.add(item.id);
    if (!families.has(item.family)) errors.push(`${prefix}: invalid family ${item.family}`);
    if (skillsByFamily.get(item.family) !== item.skill) errors.push(`${prefix}: skill ${item.skill} does not match family ${item.family}`);
    if (!statuses.has(item.status)) errors.push(`${prefix}: invalid status ${item.status}`);
    if (!outcomes.has(item.expectedOutcome)) errors.push(`${prefix}: invalid expectedOutcome ${item.expectedOutcome}`);
    if (typeof item.prompt !== "string" || item.prompt.length < 80) errors.push(`${prefix}: prompt is too short to be a realistic forward test`);
    if (!item.policy?.sourceImmutable) errors.push(`${prefix}: sourceImmutable must be true`);
    if (item.policy?.network !== false) errors.push(`${prefix}: network must be false`);
    if (!savePolicies.has(item.policy?.save)) errors.push(`${prefix}: invalid save policy ${item.policy?.save}`);
    if (!Number.isInteger(item.policy?.trials) || item.policy.trials < 1) errors.push(`${prefix}: trials must be a positive integer`);
    if (!Array.isArray(item.inputs) || !item.inputs.length) errors.push(`${prefix}: at least one declared input is required`);
    for (const input of item.inputs || []) {
      let normalized;
      try { normalized = safeRelative(input.to, `${prefix} input.to`); } catch (error) { errors.push(error.message); }
      if (!normalized || !isUnder(normalized, "inputs")) errors.push(`${prefix}: input must stay under inputs/`);
      if (inputPaths.has(normalized)) errors.push(`${prefix}: duplicate input path ${input.to}`);
      if (normalized) inputPaths.add(normalized);
      if (!new Set(["generated", "inline", "asset", "repo"]).has(input.kind)) errors.push(`${prefix}: unsupported input kind ${input.kind}`);
      if (input.kind === "generated" && !input.generator) errors.push(`${prefix}: generated input is missing generator`);
      if (input.kind === "asset" && !input.asset) errors.push(`${prefix}: asset input is missing asset path`);
    }
    for (const deliverable of item.deliverables || []) {
      let normalized;
      try { normalized = safeRelative(deliverable.path, `${prefix} deliverable.path`); } catch (error) { errors.push(error.message); }
      if (!normalized || !isUnder(normalized, "outputs")) errors.push(`${prefix}: deliverable must stay under outputs/`);
      if (deliverablePaths.has(normalized)) errors.push(`${prefix}: duplicate deliverable path ${deliverable.path}`);
      if (normalized) deliverablePaths.add(normalized);
      if (typeof deliverable.mime !== "string" || !deliverable.mime.includes("/")) errors.push(`${prefix}: deliverable ${deliverable.path} has no valid MIME type`);
    }
    if (item.expectedOutcome === "safe-refusal" && item.deliverables?.length) errors.push(`${prefix}: safe-refusal cases must not declare deliverables`);
    if (item.expectedOutcome !== "safe-refusal" && !deliverablePaths.has(path.join("outputs", "audit.json"))) errors.push(`${prefix}: successful outcomes must declare outputs/audit.json`);
    if (item.status === "ready" && item.inputs?.some((input) => input.kind === "asset")) errors.push(`${prefix}: ready cases cannot depend on locked assets`);
    if (item.status === "asset-required" && !item.inputs?.some((input) => input.kind === "asset")) errors.push(`${prefix}: asset-required case has no asset input`);
    for (const source of item.sources || []) if (!/^https:\/\//.test(source)) errors.push(`${prefix}: source must be an HTTPS URL`);
  }
  const pdfCases = cases.filter((item) => item.family === "pdf").length;
  if (pdfCases / Math.max(cases.length, 1) <= MINIMUM_PDF_CASE_SHARE) {
    errors.push(`PDF must remain an absolute majority of the suite; found ${pdfCases}/${cases.length}`);
  }
  for (const family of families) {
    const familyCases = cases.filter((item) => item.family === family);
    if (!familyCases.some((item) => item.expectedOutcome === "success")) errors.push(`${family}: missing success case`);
    if (!familyCases.some((item) => item.expectedOutcome === "safe-refusal")) errors.push(`${family}: missing safe-refusal case`);
  }
  if (errors.length) fail(errors.join("\n"));
  return { cases: cases.length, pdfCases, ready: cases.filter((item) => item.status === "ready").length };
}

function visibleCase(suite, item) {
  const inputs = (item.inputs || []).map(({ to, agentDescription }) => ({ path: to, description: agentDescription }));
  const deliverables = item.deliverables || [];
  const inputLines = inputs.map((input) => `- ${input.path}${input.description ? ` — ${input.description}` : ""}`).join("\n") || "- 无附件";
  const deliverableLines = deliverables.map((output) => `- ${output.path} (${output.mime})`).join("\n") || "- 无；本题正确结果是安全拒绝，不生成修改版 artifact。";
  return {
    id: item.id,
    family: item.family,
    skill: item.skill,
    prompt: `${suite.agentPreamble}\n\n用户任务：\n${item.prompt}\n\n可用输入：\n${inputLines}\n\n成功时必须使用以下精确交付路径：\n${deliverableLines}`,
    inputs,
    deliverables,
  };
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function oracleFingerprint(item) {
  return sha256(Buffer.from(JSON.stringify({ expectedOutcome: item.expectedOutcome, grade: item.grade }), "utf8"));
}

async function hashFile(filePath) {
  return sha256(await fs.readFile(filePath));
}

function repositoryProvenance() {
  const head = run("git", ["rev-parse", "HEAD"]).stdout.trim();
  const status = run("git", ["status", "--porcelain=v1", "--untracked-files=all"]).stdout;
  const trackedDiff = run("git", ["diff", "--binary", "HEAD", "--"]).stdout;
  return {
    head,
    dirty: status.trim().length > 0,
    statusSha256: sha256(Buffer.from(status, "utf8")),
    trackedDiffSha256: sha256(Buffer.from(trackedDiff, "utf8")),
  };
}

function run(command, args, options = {}) {
  const result = spawnSync(command, args, {
    cwd: options.cwd || repoRoot,
    encoding: options.encoding || "utf8",
    input: options.input,
    env: { ...process.env, ...options.env },
    maxBuffer: 64 * 1024 * 1024,
  });
  if (result.status !== 0) fail(`${options.label || `${command} ${args.join(" ")}`} failed (${result.status})\n${result.stdout || ""}\n${result.stderr || ""}`);
  return result;
}

const pythonGenerators = String.raw`
import json, pathlib, sys
from reportlab.lib.pagesizes import A4, landscape, letter
from reportlab.pdfgen import canvas

kind = sys.argv[1]
out = pathlib.Path(sys.argv[2])
out.parent.mkdir(parents=True, exist_ok=True)

def base_page(c, title, page, pages=1):
    width, height = c._pagesize
    c.setFont("Helvetica-Bold", 15)
    c.drawString(72, height - 48, title)
    c.setFont("Helvetica", 9)
    c.drawRightString(width - 72, 36, f"Page {page} of {pages}")

if kind == "contract-five-page":
    c = canvas.Canvas(str(out), pagesize=letter, invariant=1)
    for page in range(1, 6):
        base_page(c, "Supply Agreement", page, 5)
        c.setFont("Helvetica", 11)
        c.drawString(72, 700, f"Section {page}. This page contains ordinary contract language.")
        c.rect(72, 610, 468, 44)
        if page == 3:
            c.drawString(84, 628, "Contract ID: ACME-2025-041")
        else:
            c.drawString(84, 628, f"Reference schedule {page}")
        c.drawString(72, 570, "All non-target text and geometry must remain unchanged.")
        c.showPage()
    c.save()
elif kind == "source-bound-highlight":
    c = canvas.Canvas(str(out), pagesize=letter, invariant=1)
    for page in range(1, 4):
        base_page(c, "Governance Review Memorandum", page, 3)
        c.setFont("Helvetica", 11)
        c.drawString(72, 700, f"Review section {page}. Preserve all ordinary text and page geometry.")
        c.rect(72, 610, 468, 44)
        if page == 2:
            c.drawString(84, 628, "Conditional approval pending legal review")
        else:
            c.drawString(84, 628, f"Supporting evidence schedule {page}")
        c.drawString(72, 570, "A review marker may change only the requested source-bound selection.")
        c.showPage()
    c.save()
elif kind == "overflow-table":
    c = canvas.Canvas(str(out), pagesize=letter, invariant=1)
    base_page(c, "Approval Matrix", 1)
    c.setFont("Helvetica", 11)
    c.drawString(72, 680, "Status")
    c.rect(130, 660, 70, 24)
    c.drawString(136, 668, "Approved")
    c.drawString(72, 638, "Do not cover this adjacent line or change the table geometry.")
    c.save()
elif kind.startswith("merge-"):
    page_counts = {"merge-cover": 1, "merge-report": 2, "merge-appendix": 3}
    page_sizes = {"merge-cover": letter, "merge-report": landscape(letter), "merge-appendix": A4}
    count = page_counts[kind]
    label = kind.removeprefix("merge-").title()
    width, height = page_sizes[kind]
    c = canvas.Canvas(str(out), pagesize=(width, height), invariant=1)
    for page in range(1, count + 1):
        base_page(c, label, page, count)
        c.setFont("Helvetica", 12)
        c.drawString(72, height - 112, f"{label} source page {page}")
        c.bookmarkPage(f"{label.lower()}-{page}")
        c.addOutlineEntry(f"{label} {page}", f"{label.lower()}-{page}", 0)
        target = page % count + 1
        c.setFillColorRGB(0.05, 0.25, 0.65)
        c.drawString(72, height - 150, f"Internal link to {label} page {target}")
        c.linkRect("", f"{label.lower()}-{target}", (68, height - 156, 280, height - 132), relative=0, thickness=1)
        c.setFillColorRGB(0, 0, 0)
        c.showPage()
    c.save()
    from pypdf import PdfReader, PdfWriter
    source_reader = PdfReader(str(out), strict=True)
    named_writer = PdfWriter(clone_from=source_reader)
    for page in range(1, count + 1):
        named_writer.add_named_destination(f"{label.lower()}-named-{page}", page - 1)
    named_output = out.with_suffix(".named.pdf")
    named_writer.write(str(named_output))
    named_output.replace(out)
elif kind == "acroform-profile":
    c = canvas.Canvas(str(out), pagesize=letter, invariant=1)
    base_page(c, "Vendor Profile", 1)
    form = c.acroForm
    fields = [("full_name", "Full name", 680), ("address", "Address", 635), ("effective_date", "Date", 590), ("tin", "TIN (leave blank)", 545), ("signature", "Signature (leave blank)", 500)]
    for name, label, y in fields:
        c.setFont("Helvetica", 10); c.drawString(72, y + 7, label)
        form.textfield(name=name, tooltip=label, x=190, y=y, width=260, height=24, borderWidth=1)
    c.drawString(72, 460, "Company type")
    form.radio(name="company_type", value="LLC", selected=False, x=190, y=455, buttonStyle="circle")
    c.drawString(210, 460, "LLC")
    form.radio(name="company_type", value="Corporation", selected=False, x=260, y=455, buttonStyle="circle")
    c.drawString(280, 460, "Corporation")
    c.drawString(72, 415, "Existing acknowledgment (preserve)")
    form.checkbox(name="terms_ack", checked=True, x=260, y=410, buttonStyle="check")
    c.save()
elif kind in {"attachment-portfolio", "active-content"}:
    from pypdf import PdfReader, PdfWriter
    from pypdf.annotations import Text
    from pypdf.generic import DecodedStreamObject, DictionaryObject, NameObject, NumberObject, RectangleObject, TextStringObject, create_string_object
    temporary = out.with_suffix(".base.pdf")
    c = canvas.Canvas(str(temporary), pagesize=letter, invariant=1)
    base_page(c, "Public Release Review", 1)
    c.setFont("Helvetica", 11)
    c.drawString(72, 680, "This visible page content must remain stable.")
    if kind == "active-content":
        c.acroForm.textfield(name="reviewer", tooltip="Reviewer", x=72, y=620, width=180, height=24, value="Private Person")
        hidden = c.beginText(72, 580)
        hidden.setTextRenderMode(3)
        hidden.textLine("HIDDEN-REVIEW-CANARY-7A91")
        c.drawText(hidden)
    c.save()
    reader = PdfReader(str(temporary))
    writer = PdfWriter(clone_from=reader)
    writer.add_metadata({"/Author": "Private Person", "/Subject": "Internal only"})
    attachments = [("report.txt", b"first"), ("report.txt", b"second"), ("unicode-\u6d4b\u8bd5.txt", b"unicode"), ("../escape.exe", b"MZ-not-executable"), ("archive.zip", b"PK-not-opened")]
    if kind == "active-content":
        attachments[0] = ("internal-review.txt", b"ATTACHMENT-CANARY-4D27")
    mime_types = {".txt": "text/plain", ".zip": "application/zip", ".exe": "application/vnd.microsoft.portable-executable"}
    for name, payload in attachments:
        embedded = writer.add_attachment(name, payload)
        suffix = pathlib.Path(name.replace("\\\\", "/")).suffix.lower()
        embedded.subtype = NameObject("/" + mime_types.get(suffix, "application/octet-stream").replace("/", "#2F"))
    if kind == "attachment-portfolio":
        page_payload = b"page-level-review"
        page_stream = DecodedStreamObject()
        page_stream.set_data(page_payload)
        page_stream[NameObject("/Type")] = NameObject("/EmbeddedFile")
        page_stream[NameObject("/Subtype")] = NameObject("/text#2Fplain")
        page_stream_reference = writer._add_object(page_stream)
        page_filespec = DictionaryObject({
            NameObject("/Type"): NameObject("/Filespec"),
            NameObject("/F"): TextStringObject("report.txt"),
            NameObject("/UF"): TextStringObject("report.txt"),
            NameObject("/Desc"): TextStringObject("Page-level reviewer attachment"),
            NameObject("/EF"): DictionaryObject({NameObject("/F"): page_stream_reference, NameObject("/UF"): page_stream_reference}),
        })
        page_annotation = DictionaryObject({
            NameObject("/Type"): NameObject("/Annot"),
            NameObject("/Subtype"): NameObject("/FileAttachment"),
            NameObject("/Rect"): RectangleObject((500, 680, 520, 700)),
            NameObject("/FS"): writer._add_object(page_filespec),
            NameObject("/Contents"): TextStringObject("Page-level reviewer attachment"),
            NameObject("/Name"): NameObject("/PushPin"),
            NameObject("/F"): NumberObject(4),
        })
        writer.add_annotation(0, page_annotation)
    if kind == "active-content":
        writer.add_js("app.alert('JS-CANARY-2F61');")
        writer._root_object["/Names"]["/JavaScript"]["/Names"][0] = create_string_object("00000000-0000-0000-0000-000000000001")
        writer._root_object[NameObject("/OpenAction")] = DictionaryObject({
            NameObject("/S"): NameObject("/JavaScript"),
            NameObject("/JS"): TextStringObject("app.alert('OPENACTION-CANARY-8B03');"),
        })
        writer._root_object[NameObject("/AA")] = DictionaryObject({
            NameObject("/WC"): DictionaryObject({
                NameObject("/S"): NameObject("/Launch"),
                NameObject("/F"): TextStringObject("LAUNCH-CANARY-6C42.exe"),
            }),
            NameObject("/WP"): DictionaryObject({
                NameObject("/S"): NameObject("/SubmitForm"),
                NameObject("/F"): TextStringObject("https://invalid.example/SUBMIT-CANARY-9E18"),
            }),
        })
        writer.add_annotation(0, Text(rect=(500, 680, 520, 700), text="COMMENT-CANARY-3A55", open=False))
    with out.open("wb") as handle: writer.write(handle)
    temporary.unlink()
else:
    raise SystemExit(f"unknown generator: {kind}")
`;

async function generateInput(generator, target) {
  const officeFixture = await generateOfficeInput(generator, target);
  if (officeFixture) return officeFixture;
  const python = generatorPythonExecutable();
  run(python, ["-c", pythonGenerators, generator, target], { label: `generate fixture ${generator}` });
}

let generatorPython;
function generatorPythonExecutable() {
  if (generatorPython) return generatorPython;
  const python = process.env.OPEN_OFFICE_AGENT_EVAL_PYTHON || process.env.OPEN_OFFICE_PDF_PROVIDER_PYTHON || "python3";
  const probe = spawnSync(python, ["-c", "import reportlab, pypdf"], { encoding: "utf8", env: process.env });
  if (probe.status !== 0) {
    fail(`Generated Agent eval fixtures require a Python environment with reportlab and pypdf. Set OPEN_OFFICE_AGENT_EVAL_PYTHON to that interpreter (for Codex, use the Python path returned by load_workspace_dependencies). Probe failed for ${python}: ${String(probe.stderr || probe.error?.message || "unknown error").trim()}`);
  }
  generatorPython = python;
  return generatorPython;
}

async function copyTree(source, target) {
  await fs.cp(source, target, { recursive: true, force: true });
}

async function copySkillTree(source, target) {
  await fs.cp(source, target, {
    recursive: true,
    force: true,
    filter: (entry) => path.basename(entry) !== "__pycache__" && path.extname(entry).toLowerCase() !== ".pyc",
  });
}

async function patchReferenceSkillPackageName(root) {
  const changed = [];
  const textExtensions = new Set([".md", ".mjs", ".js", ".cjs", ".json", ".yaml", ".yml", ".py", ".txt"]);
  async function walk(directory) {
    for (const entry of await fs.readdir(directory, { withFileTypes: true })) {
      const target = path.join(directory, entry.name);
      if (entry.isDirectory()) await walk(target);
      else if (entry.isFile() && textExtensions.has(path.extname(entry.name).toLowerCase())) {
        const source = await fs.readFile(target, "utf8");
        const patched = source.replaceAll(/(?<!open-)office-artifact-tool/g, "open-office-artifact-tool");
        if (patched !== source) {
          await fs.writeFile(target, patched, "utf8");
          changed.push(path.relative(root, target));
        }
      }
    }
  }
  await walk(root);
  return changed.sort();
}

export function skillSource(item, subject) {
  const skillName = item.skill;
  if (subject === "candidate") return path.join(repoRoot, "skills", item.family, "skills", skillName);
  if (subject === "reference") return path.join(REFERENCE_SKILLS_ROOT, item.family, "skills", skillName);
  fail(`unknown subject ${subject}`);
}

async function materializeInput(input, workspace) {
  const target = path.join(workspace, safeRelative(input.to, "input.to"));
  await fs.mkdir(path.dirname(target), { recursive: true });
  if (input.kind === "inline") await fs.writeFile(target, input.content, "utf8");
  else if (input.kind === "generated") await generateInput(input.generator, target);
  else if (input.kind === "repo") await copyTree(path.join(repoRoot, safeRelative(input.source, "input.source")), target);
  else if (input.kind === "asset") {
    const source = path.join(repoRoot, "evals", "assets", safeRelative(input.asset, "input.asset"));
    try { await fs.access(source); } catch { fail(`missing locked eval asset: ${path.relative(repoRoot, source)}\nFixture requirement: ${input.fixtureSpec || "not documented"}`); }
    await copyTree(source, target);
  }
  await makeReadOnly(target);
}

async function makeReadOnly(target) {
  const stat = await fs.lstat(target);
  if (stat.isSymbolicLink()) fail(`eval inputs cannot contain symbolic links: ${target}`);
  if (stat.isDirectory()) {
    for (const entry of await fs.readdir(target)) await makeReadOnly(path.join(target, entry));
    await fs.chmod(target, 0o555);
    return;
  }
  if (!stat.isFile()) fail(`eval input must be a regular file or directory: ${target}`);
  await fs.chmod(target, 0o444);
}

async function removePreparedTree(target) {
  let stat;
  try { stat = await fs.lstat(target); } catch { return; }
  if (stat.isDirectory()) {
    await fs.chmod(target, 0o700);
    for (const entry of await fs.readdir(target)) await removePreparedTree(path.join(target, entry));
  } else if (stat.isFile()) {
    await fs.chmod(target, 0o600);
  }
  await fs.rm(target, { recursive: true, force: true });
}

async function fingerprintPath(target) {
  let stat;
  try { stat = await fs.lstat(target); } catch { return null; }
  if (stat.isSymbolicLink()) fail(`cannot fingerprint symbolic link input: ${target}`);
  if (stat.isFile()) return hashFile(target);
  if (stat.isDirectory()) return `tree:${await hashTree(target)}`;
  fail(`cannot fingerprint non-file input: ${target}`);
}

async function packageCandidate(evaluatorDir, workspace) {
  const packDir = path.join(evaluatorDir, "package");
  await fs.mkdir(packDir, { recursive: true });
  const packed = run("npm", ["pack", "--json", "--ignore-scripts", "--pack-destination", packDir], { cwd: repoRoot });
  const report = JSON.parse(packed.stdout)[0];
  const tarball = path.join(packDir, report.filename);
  await fs.writeFile(path.join(workspace, "package.json"), JSON.stringify({ name: "agent-eval-workspace", private: true, type: "module" }, null, 2));
  run("npm", ["install", "--ignore-scripts", "--no-audit", "--no-fund", tarball], { cwd: workspace });
  return { tarball: path.relative(evaluatorDir, tarball), sha256: await hashFile(tarball) };
}

export function providerRuntimeInstruction(item, environment = process.env) {
  const configured = environment.OPEN_OFFICE_AGENT_EVAL_PYTHON || environment.OPEN_OFFICE_PDF_PROVIDER_PYTHON;
  if (item?.family !== "pdf" || !configured) return "";
  const interpreter = path.resolve(configured);
  return `\n## Prepared provider runtime\n\nThis isolated trial supplies the authoritative PDF provider interpreter below. Export it before every PDF Python probe, plan, mutation, scan, and audit command. Do not replace it, install another environment, or fall back to system Python.\n\n\`\`\`bash\nexport OPEN_OFFICE_PDF_PROVIDER_PYTHON=${JSON.stringify(interpreter)}\n\`\`\`\n`;
}

async function prepareCase(suite, item, options) {
  if (item.status !== "ready") fail(`${item.id} is asset-required; add pinned files under evals/assets before preparing it`);
  const subject = options.subject || "candidate";
  const trial = Number(options.trial || 1);
  if (!Number.isInteger(trial) || trial < 1) fail("--trial must be a positive integer");
  const runRoot = path.resolve(options.runRoot || path.join(os.tmpdir(), "open-office-agent-evals", new Date().toISOString().replace(/[:.]/g, "-")));
  const trialRoot = path.join(runRoot, item.id, `${subject}-trial-${trial}`);
  const workspace = path.join(trialRoot, "workspace");
  const evaluator = path.join(trialRoot, "evaluator");
  await removePreparedTree(trialRoot);
  await fs.mkdir(path.join(workspace, "outputs"), { recursive: true });
  await fs.mkdir(evaluator, { recursive: true });
  if (subject === "reference") await checkReferenceSkillSync();
  const sourceSkill = skillSource(item, subject);
  const sourceSkillStat = await fs.stat(sourceSkill).catch(() => null);
  if (!sourceSkillStat?.isDirectory()) {
    fail(`missing ${subject} Skill source: ${path.relative(repoRoot, sourceSkill)}${subject === "reference" ? " (initialize the pinned reference submodule first)" : ""}`);
  }
  const installedSkill = path.join(workspace, ".agents", "skills", item.skill);
  await copySkillTree(sourceSkill, installedSkill);
  const referencePackageNamePatches = subject === "reference" ? await patchReferenceSkillPackageName(installedSkill) : [];
  await makeReadOnly(installedSkill);
  for (const input of item.inputs || []) await materializeInput(input, workspace);
  const inputHashes = {};
  for (const input of item.inputs || []) {
    const target = path.join(workspace, safeRelative(input.to, "input.to"));
    inputHashes[input.to] = await fingerprintPath(target);
  }
  const packageRecord = await packageCandidate(evaluator, workspace);
  const prompt = `${visibleCase(suite, item).prompt}${providerRuntimeInstruction(item)}`;
  await fs.writeFile(path.join(workspace, "PROMPT.md"), `${prompt}\n`, "utf8");
  const workspaceHashes = {};
  for (const relative of ["PROMPT.md", "package.json", "package-lock.json", ".agents", "node_modules"]) {
    workspaceHashes[relative] = await fingerprintPath(path.join(workspace, relative));
  }
  const repository = repositoryProvenance();
  const runRecord = {
    suite: suite.id,
    case: item.id,
    subject,
    trial,
    preparedAt: new Date().toISOString(),
    git: repository.head,
    gitWorktree: {
      dirty: repository.dirty,
      statusSha256: repository.statusSha256,
      trackedDiffSha256: repository.trackedDiffSha256,
    },
    package: packageRecord,
    referencePackageNamePatches,
    skillSha256: await hashTree(installedSkill),
    workspaceHashes,
    inputHashes,
    oracleSha256: oracleFingerprint(item),
  };
  await fs.writeFile(path.join(evaluator, "run.json"), JSON.stringify(runRecord, null, 2), "utf8");
  return { trialRoot, workspace, evaluator, promptPath: path.join(workspace, "PROMPT.md") };
}

async function hashTree(root) {
  const entries = [];
  async function walk(directory) {
    for (const entry of await fs.readdir(directory, { withFileTypes: true })) {
      const target = path.join(directory, entry.name);
      const mode = (await fs.lstat(target)).mode & 0o777;
      if (entry.isDirectory()) {
        entries.push({ target, type: "directory", mode });
        await walk(target);
      }
      else if (entry.isFile()) entries.push({ target, type: "file", mode });
      else if (entry.isSymbolicLink()) entries.push({ target, type: "symlink", mode });
      else fail(`hashed trees can contain only regular files and directories: ${target}`);
    }
  }
  await walk(root);
  entries.sort((left, right) => left.target.localeCompare(right.target));
  const digest = crypto.createHash("sha256");
  digest.update(`root\0${((await fs.lstat(root)).mode & 0o777).toString(8)}\0`);
  for (const entry of entries) {
    digest.update(`${entry.type}\0`);
    digest.update(`${entry.mode.toString(8)}\0`);
    digest.update(path.relative(root, entry.target).split(path.sep).join("/"));
    digest.update("\0");
    if (entry.type === "file") digest.update(await fs.readFile(entry.target));
    else if (entry.type === "symlink") digest.update(await fs.readlink(entry.target));
    digest.update("\0");
  }
  return digest.digest("hex");
}

async function runCodex(prepared, options) {
  const prompt = await fs.readFile(prepared.promptPath, "utf8");
  const args = ["exec", "--ephemeral", "--ignore-user-config", "--ignore-rules", "--json", "--sandbox", "workspace-write", "--skip-git-repo-check", "-C", prepared.workspace, "-o", path.join(prepared.evaluator, "final.txt")];
  if (options.model) args.push("--model", options.model);
  args.push("-");
  const result = spawnSync(options.codex || "codex", args, {
    cwd: prepared.workspace,
    encoding: "utf8",
    input: prompt,
    env: { ...process.env, PYTHONDONTWRITEBYTECODE: "1" },
    maxBuffer: 256 * 1024 * 1024,
  });
  await fs.writeFile(path.join(prepared.evaluator, "trace.jsonl"), result.stdout || "", "utf8");
  await fs.writeFile(path.join(prepared.evaluator, "stderr.txt"), result.stderr || "", "utf8");
  await fs.writeFile(path.join(prepared.evaluator, "exit.json"), JSON.stringify({ status: result.status, signal: result.signal }, null, 2), "utf8");
  return result.status;
}

async function inventoryRelativeFiles(root) {
  const files = [];
  const invalid = [];
  async function walk(directory) {
    for (const entry of await fs.readdir(directory, { withFileTypes: true })) {
      const target = path.join(directory, entry.name);
      if (entry.isDirectory()) await walk(target);
      else if (entry.isFile()) files.push(path.relative(root, target).split(path.sep).join("/"));
      else invalid.push({ path: path.relative(root, target).split(path.sep).join("/"), type: entry.isSymbolicLink() ? "symlink" : "non-regular" });
    }
  }
  try { await walk(root); } catch (error) { if (error.code !== "ENOENT") throw error; }
  return { files: files.sort(), invalid: invalid.sort((left, right) => left.path.localeCompare(right.path)) };
}

async function scorePrepared(item, prepared, options = {}) {
  const runRecord = JSON.parse(await fs.readFile(path.join(prepared.evaluator, "run.json"), "utf8"));
  const checks = [];
  checks.push({ id: "oracle-version", gate: true, passed: runRecord.oracleSha256 === oracleFingerprint(item), expected: runRecord.oracleSha256, actual: oracleFingerprint(item) });
  const exitRecord = await fs.readFile(path.join(prepared.evaluator, "exit.json"), "utf8").then(JSON.parse, () => null);
  const finalMessage = await fs.readFile(path.join(prepared.evaluator, "final.txt"), "utf8").catch(() => "");
  const trace = await fs.readFile(path.join(prepared.evaluator, "trace.jsonl"), "utf8").catch(() => "");
  checks.push({ id: "execution-complete", gate: true, passed: exitRecord?.status === 0 && Boolean(finalMessage.trim()) && Boolean(trace.trim()), actual: { status: exitRecord?.status ?? null, finalMessage: Boolean(finalMessage.trim()), trace: Boolean(trace.trim()) } });
  for (const [relative, before] of Object.entries(runRecord.workspaceHashes || {})) {
    const after = await fingerprintPath(path.join(prepared.workspace, relative));
    checks.push({ id: `workspace-immutable:${relative}`, gate: true, passed: before === after, expected: before, actual: after });
  }
  for (const [relative, before] of Object.entries(runRecord.inputHashes)) {
    const after = await fingerprintPath(path.join(prepared.workspace, relative));
    checks.push({ id: `source-immutable:${relative}`, gate: true, passed: before === after, expected: before, actual: after });
  }
  const outputInventory = await inventoryRelativeFiles(path.join(prepared.workspace, "outputs"));
  const outputEntries = outputInventory.files;
  checks.push({ id: "output-regular-files-only", gate: true, passed: outputInventory.invalid.length === 0, actual: outputInventory.invalid });
  const nonAuditOutputs = outputEntries.filter((entry) => entry !== "audit.json");
  const observedOutcome = item.expectedOutcome === "safe-refusal" || item.expectedOutcome === "success-or-safe-refusal" && nonAuditOutputs.length === 0
    ? "safe-refusal"
    : "success";
  if (observedOutcome === "safe-refusal") {
    const unexpected = outputEntries.filter((entry) => entry !== "audit.json");
    checks.push({ id: "no-artifact-on-safe-refusal", gate: true, passed: unexpected.length === 0, actual: outputEntries });
    if (outputEntries.includes("audit.json")) {
      let audit;
      try { audit = JSON.parse(await fs.readFile(path.join(prepared.workspace, "outputs", "audit.json"), "utf8")); } catch { audit = null; }
      checks.push({ id: "safe-refusal-audit-status", gate: true, passed: audit?.status === "failed_closed" && audit?.delivered_modified_pdf !== true, actual: audit?.status || "invalid-json" });
    }
    const diagnosticTerms = item.grade?.machine?.diagnosticTerms || [];
    if (diagnosticTerms.length) checks.push({ id: "safe-refusal-diagnostic", gate: false, passed: diagnosticTerms.some((term) => finalMessage.toLowerCase().includes(String(term).toLowerCase())), expectedAny: diagnosticTerms });
  } else {
    const declaredFiles = new Set((item.deliverables || [])
      .filter((deliverable) => deliverable.mime !== "inode/directory")
      .map((deliverable) => path.relative("outputs", safeRelative(deliverable.path, "deliverable.path")).split(path.sep).join("/")));
    const declaredDirectories = (item.deliverables || [])
      .filter((deliverable) => deliverable.mime === "inode/directory")
      .map((deliverable) => path.relative("outputs", safeRelative(deliverable.path, "deliverable.path")).split(path.sep).join("/"));
    const declaredOutputs = [...declaredFiles, ...declaredDirectories].sort();
    checks.push({
      id: "no-undeclared-deliverables",
      gate: true,
      passed: outputEntries.every((entry) => declaredFiles.has(entry) || declaredDirectories.some((directory) => entry.startsWith(`${directory}/`))),
      expected: declaredOutputs,
      actual: outputEntries,
    });
    for (const deliverable of item.deliverables || []) {
      const target = path.join(prepared.workspace, safeRelative(deliverable.path, "deliverable.path"));
      if (deliverable.mime === "inode/directory") {
        let directoryInventory = null;
        try {
          if ((await fs.stat(target)).isDirectory()) directoryInventory = await inventoryRelativeFiles(target);
        } catch {}
        checks.push({
          id: `deliverable:${deliverable.path}`,
          gate: true,
          passed: Boolean(directoryInventory && directoryInventory.files.length && directoryInventory.invalid.length === 0),
          actual: directoryInventory,
        });
        continue;
      }
      let bytes;
      try { bytes = await fs.readFile(target); } catch { bytes = null; }
      checks.push({ id: `deliverable:${deliverable.path}`, gate: true, passed: Boolean(bytes?.length), actualBytes: bytes?.length || 0 });
      const magic = mimeMagic.get(deliverable.mime);
      if (bytes && magic) checks.push({ id: `magic:${deliverable.path}`, gate: true, passed: bytes.subarray(0, magic.length).equals(magic) });
      if (bytes && deliverable.mime === "application/json") {
        let parsed = false;
        try { JSON.parse(bytes.toString("utf8")); parsed = true; } catch {}
        checks.push({ id: `json:${deliverable.path}`, gate: true, passed: parsed });
      }
    }
  }
  const caseGrade = await (item.family === "pdf" ? gradePdfCase : gradeOfficeCase)({
    item,
    workspace: prepared.workspace,
    evaluator: prepared.evaluator,
    finalMessage,
    trace,
    weights: options.weights,
  });
  if (caseGrade.supported) checks.push(...caseGrade.checks);
  const hardGatesPassed = checks.filter((check) => check.gate).every((check) => check.passed);
  const caseSpecificPassed = caseGrade.supported && caseGrade.graded ? caseGrade.caseSpecificPassed : null;
  const report = {
    case: item.id,
    expectedOutcome: item.expectedOutcome,
    observedOutcome,
    checks,
    hardGatesPassed,
    taskPassed: caseSpecificPassed === null ? null : hardGatesPassed && caseSpecificPassed,
    machineScoreStatus: caseGrade.supported
      ? caseGrade.graded
        ? hardGatesPassed && caseSpecificPassed ? "case-graded-pass" : "case-graded-fail"
        : "grader-unavailable"
      : item.expectedOutcome === "safe-refusal" ? "generic-refusal-gates" : "partial-generic-only",
    pending: caseGrade.supported
      ? caseGrade.pending
      : ["case-specific semantic oracle", "all-page visual grading", "security/provider trace grading"],
  };
  if (caseGrade.supported) {
    report.categoryScores = caseGrade.categoryScores || null;
    report.rawScorePercent = caseGrade.rawScorePercent ?? null;
    report.scorePercent = caseGrade.graded ? hardGatesPassed ? caseGrade.rawScorePercent : 0 : null;
    report.graderEvidence = caseGrade.evidence;
    if (caseGrade.infrastructureErrors?.length) report.infrastructureErrors = caseGrade.infrastructureErrors;
  }
  await fs.writeFile(path.join(prepared.evaluator, "report.json"), JSON.stringify(report, null, 2), "utf8");
  return report;
}

function help() {
  return `Agent artifact PromptBench\n\nCommands:\n  validate\n  list [--family pdf] [--status ready] [--json]\n  show <case-id> [--json]\n  prepare <case-id> [--subject candidate|reference] [--trial 1] [--run-root <path>]\n  run <case-id> [prepare options] [--model <model>] [--codex <path>]\n  score <case-id> --trial-root <prepared trial directory>\n\nThe Agent receives only PROMPT.md, declared inputs, the selected Skill, and an installed candidate tarball. Prepared PDF trials also declare one authoritative provider interpreter in PROMPT.md when the runner was given OPEN_OFFICE_AGENT_EVAL_PYTHON or OPEN_OFFICE_PDF_PROVIDER_PYTHON; the same interpreter is used for generated fixtures and hidden oracles. Fixture specifications and graders are never copied into its workspace; run.json records integrity evidence and only a fingerprint of the hidden oracle, never the grading specification. The default run root is outside the repository in the OS temp directory. A production benchmark must additionally mount only the trial workspace into a no-network container, because a CLI sandbox alone is not an oracle confidentiality boundary. The eight ready PDF cases plus the ready XLSX threaded-comment, growth-assumption, connection refresh-on-open, and Pivot refresh-on-open cases, two DOCX cases (classic-comment and source-bound DOCX header text), and three PPTX cases (title plus fixed-topology rich-notes run edit, source-bound slide-name edit, and closed-leaf slide clone) have independent semantic, native-render, security, and provider-trace graders. PDF must remain an absolute majority of the full suite; the remaining 21 asset-required cases never claim full success before pinned corpus or PKI fixtures exist.\n`;
}

export async function main(argv = process.argv.slice(2)) {
  const { command, positionals, flags } = parseArgs(argv);
  const { suite, cases } = await loadSuite();
  if (command === "help" || command === "--help" || command === "-h") console.log(help());
  else if (command === "validate") console.log(JSON.stringify({ ok: true, suite: suite.id, ...validateSuite(suite, cases) }, null, 2));
  else if (command === "list") {
    validateSuite(suite, cases);
    const filtered = cases.filter((item) => (!flags.get("family") || item.family === flags.get("family")) && (!flags.get("status") || item.status === flags.get("status")));
    if (flags.get("json")) console.log(JSON.stringify(filtered.map(({ id, family, status, expectedOutcome, tags }) => ({ id, family, status, expectedOutcome, tags })), null, 2));
    else for (const item of filtered) console.log([item.id, item.family, item.status, item.expectedOutcome].join("\t"));
  } else if (command === "show") {
    validateSuite(suite, cases);
    const item = cases.find((candidate) => candidate.id === positionals[0]);
    if (!item) fail(`unknown case: ${positionals[0]}`);
    const visible = visibleCase(suite, item);
    console.log(flags.get("json") ? JSON.stringify(visible, null, 2) : visible.prompt);
  } else if (command === "prepare" || command === "run") {
    validateSuite(suite, cases);
    const item = cases.find((candidate) => candidate.id === positionals[0]);
    if (!item) fail(`unknown case: ${positionals[0]}`);
    const prepared = await prepareCase(suite, item, { subject: flags.get("subject"), trial: flags.get("trial"), runRoot: flags.get("run-root") });
    let exitStatus = null;
    let report = null;
    if (command === "run") {
      exitStatus = await runCodex(prepared, { model: flags.get("model"), codex: flags.get("codex") });
      report = await scorePrepared(item, prepared, { weights: suite.weights });
    }
    console.log(JSON.stringify({ prepared, exitStatus, report }, null, 2));
    if (command === "run" && exitStatus !== 0) process.exitCode = exitStatus || 1;
  } else if (command === "score") {
    validateSuite(suite, cases);
    const item = cases.find((candidate) => candidate.id === positionals[0]);
    if (!item) fail(`unknown case: ${positionals[0]}`);
    const trialRootFlag = flags.get("trial-root");
    if (!trialRootFlag) fail("score requires --trial-root");
    const trialRoot = path.resolve(trialRootFlag);
    const prepared = { trialRoot, workspace: path.join(trialRoot, "workspace"), evaluator: path.join(trialRoot, "evaluator"), promptPath: path.join(trialRoot, "workspace", "PROMPT.md") };
    console.log(JSON.stringify(await scorePrepared(item, prepared, { weights: suite.weights }), null, 2));
  } else fail(`unknown command ${command}\n\n${help()}`);
}

export { fingerprintPath, generateInput, loadSuite, makeReadOnly, oracleFingerprint, removePreparedTree, repositoryProvenance, scorePrepared, validateSuite, visibleCase };

const entryPath = process.argv[1] ? path.resolve(process.argv[1]) : "";
if (entryPath === fileURLToPath(import.meta.url)) await main();
