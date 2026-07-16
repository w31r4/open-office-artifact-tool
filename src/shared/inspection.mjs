export function ndjson(records, maxChars = Infinity) {
  const lines = records.map((record) => JSON.stringify(record));
  let text = lines.join("\n");
  let truncated = false;
  if (text.length > maxChars) {
    truncated = true;
    const kept = [];
    let chars = 0;
    for (const line of lines) {
      if (chars + line.length + 1 > maxChars) break;
      kept.push(line);
      chars += line.length + 1;
    }
    kept.push(JSON.stringify({ kind: "notice", message: `Truncated: omitted ${lines.length - kept.length} lines. Increase maxChars or narrow query.` }));
    text = kept.join("\n");
  }
  return { ndjson: text, truncated };
}

export function inspectTargetTokens(options = {}) {
  const raw = options.target ?? options.targetId ?? options.id ?? options.anchor;
  if (raw == null || raw === "") return [];
  const values = Array.isArray(raw) ? raw : [raw];
  const out = [];
  for (const value of values) {
    if (value == null) continue;
    if (typeof value === "object") {
      out.push(value.id, value.targetId, value.name, value.address, value.range, value.ref, value.sheetName && value.address ? `${value.sheetName}!${value.address}` : undefined);
    } else {
      out.push(...String(value).split(","));
    }
  }
  return out.map((value) => String(value ?? "").trim()).filter(Boolean);
}

export function inspectRecordMatchesTarget(record, targets) {
  if (!targets.length) return true;
  if (!record) return false;
  const candidates = new Set();
  const add = (value) => { if (value != null && value !== "") candidates.add(String(value)); };
  for (const key of ["id", "targetId", "parentId", "layoutId", "name", "address", "range", "sheet", "slide", "page", "kind", "drawingType", "regionKind"]) add(record[key]);
  if (record.sheet && record.address) add(`${record.sheet}!${record.address}`);
  if (record.sheet && record.range) add(`${record.sheet}!${record.range}`);
  if (record.target) {
    add(record.target.id);
    add(record.target.address);
    add(record.target.range);
    if (record.target.sheetName && record.target.address) add(`${record.target.sheetName}!${record.target.address}`);
  }
  const haystack = JSON.stringify(record);
  return targets.some((target) => candidates.has(target) || haystack.includes(target));
}

export function filterInspectRecords(records, options = {}) {
  const search = String(options.search || options.searchTerm || "").trim().toLowerCase();
  const targets = inspectTargetTokens(options);
  let filtered = records
    .filter(Boolean)
    .filter((record) => !search || JSON.stringify(record).toLowerCase().includes(search));
  if (targets.length) {
    const before = Math.max(0, Number(options.before ?? options.contextBefore ?? options.context ?? 0) || 0);
    const after = Math.max(0, Number(options.after ?? options.contextAfter ?? options.context ?? 0) || 0);
    if (before || after) {
      const keep = new Set();
      filtered.forEach((record, index) => {
        if (!inspectRecordMatchesTarget(record, targets)) return;
        for (let i = Math.max(0, index - before); i <= Math.min(filtered.length - 1, index + after); i += 1) keep.add(i);
      });
      filtered = filtered.filter((_, index) => keep.has(index));
    } else {
      filtered = filtered.filter((record) => inspectRecordMatchesTarget(record, targets));
    }
  }
  return shapeInspectRecords(filtered, options);
}

const INSPECT_CORE_FIELDS = new Set(["kind", "id", "sheet", "address", "range", "name", "page", "slide", "targetId", "parentId"]);
const INSPECT_FIELD_ALIASES = {
  values: ["values", "value"],
  value: ["value", "values"],
  formulas: ["formulas", "formula"],
  formula: ["formula", "formulas"],
  bbox: ["bbox", "bboxUnit"],
  text: ["text", "textPreview", "textChars"],
  style: ["style", "styleId"],
};

function inspectFieldList(value) {
  if (value == null || value === "") return [];
  const values = Array.isArray(value) ? value : String(value).split(",");
  return values.map((item) => String(item || "").trim()).filter(Boolean);
}

function expandedInspectFields(fields) {
  const out = new Set();
  for (const field of fields) {
    out.add(field);
    for (const alias of INSPECT_FIELD_ALIASES[field] || []) out.add(alias);
  }
  return out;
}

function shapeInspectRecord(record, options = {}) {
  const includeFields = expandedInspectFields(inspectFieldList(options.fields ?? options.includeFields ?? options.include));
  const excludeFields = expandedInspectFields(inspectFieldList(options.exclude ?? options.omit));
  if (!includeFields.size && !excludeFields.size) return record;
  const shaped = {};
  for (const [key, value] of Object.entries(record)) {
    const keepByInclude = !includeFields.size || includeFields.has(key) || INSPECT_CORE_FIELDS.has(key);
    const dropByExclude = excludeFields.has(key) && !INSPECT_CORE_FIELDS.has(key);
    if (keepByInclude && !dropByExclude) shaped[key] = value;
  }
  return shaped;
}

function shapeInspectRecords(records, options = {}) {
  return records.map((record) => shapeInspectRecord(record, options));
}

export function verificationResult(artifactKind, issues, options = {}) {
  const result = {
    artifactKind,
    ok: issues.length === 0,
    issues,
    ...ndjson(issues, options.maxChars ?? Infinity),
  };
  return result;
}

export function verificationIssue(artifactKind, type, message, details = {}) {
  return { kind: "verificationIssue", artifactKind, type, severity: details.severity || "error", message, ...details };
}


export function normalizeKinds(kind, fallback) {
  if (!kind) return new Set(fallback);
  return new Set(String(kind).split(",").map((value) => value.trim()).filter(Boolean));
}
