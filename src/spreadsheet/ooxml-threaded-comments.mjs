import path from "node:path";

import { attrEscape, attributes, decodeXml, rootTag } from "../ooxml/source-reference-xml.mjs";

export const XLSX_PERSON_CONTENT_TYPE = "application/vnd.ms-excel.person+xml";
export const XLSX_PERSON_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2017/10/relationships/person";
export const XLSX_THREADED_COMMENTS_CONTENT_TYPE = "application/vnd.ms-excel.threadedcomments+xml";
export const XLSX_THREADED_COMMENTS_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";
export const XLSX_THREADED_COMMENTS_NAMESPACE = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";

const GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/;
const decoder = new TextDecoder();

function localAttribute(tag, localName) {
  return Object.entries(attributes(tag)).find(([name]) => name === localName || name.endsWith(`:${localName}`))?.[1];
}

function normalizeGuid(value, label) {
  const guid = String(value || "").trim().toUpperCase();
  if (!GUID.test(guid)) throw new TypeError(`${label} must be a brace-delimited GUID.`);
  return guid;
}

function hashWords(value) {
  let first = 0x811C9DC5;
  let second = 0x9E3779B9;
  for (const character of String(value)) {
    const code = character.codePointAt(0);
    first = Math.imul(first ^ code, 0x01000193) >>> 0;
    second = Math.imul(second ^ (code + first), 0x85EBCA6B) >>> 0;
  }
  return [first, second, (first ^ Math.imul(second, 33)) >>> 0, (second ^ Math.imul(first, 97)) >>> 0];
}

export function deterministicSpreadsheetGuid(seed) {
  const words = hashWords(seed).map((word) => word.toString(16).toUpperCase().padStart(8, "0")).join("");
  return `{${words.slice(0, 8)}-${words.slice(8, 12)}-4${words.slice(13, 16)}-8${words.slice(17, 20)}-${words.slice(20, 32)}}`;
}

function commentPerson(comment, thread) {
  const person = comment.person || {};
  const displayName = person.displayName || comment.author || thread.author || "User";
  const id = normalizeGuid(comment.personId || person.id || deterministicSpreadsheetGuid(`person:${displayName}`), `Spreadsheet comment person ${displayName} id`);
  return { id, displayName: String(displayName), userId: String(person.userId ?? comment.userId ?? displayName), providerId: String(person.providerId ?? comment.providerId ?? "None") };
}

function commentDate(comment) {
  const value = comment.date || comment.dT || "1970-01-01T00:00:00.000Z";
  if (Number.isNaN(Date.parse(value))) throw new TypeError(`Spreadsheet threaded comment date ${value} must be valid.`);
  return new Date(value).toISOString();
}

export function planSpreadsheetThreadedComments(threadParts = []) {
  const people = new Map();
  const ids = new Set();
  const parts = threadParts.map((part) => {
    const entries = [];
    for (const [threadIndex, thread] of part.threads.entries()) {
      if (!thread.comments?.length) continue;
      const ref = String(thread.target?.address || "").trim().toUpperCase();
      if (!ref) throw new TypeError(`Spreadsheet comment thread ${thread.id || threadIndex} requires a target address.`);
      let rootId;
      for (const [commentIndex, comment] of thread.comments.entries()) {
        const id = normalizeGuid(comment.id || deterministicSpreadsheetGuid(`comment:${thread.id || `${part.sheetIndex}:${threadIndex}`}:${commentIndex}`), `Spreadsheet threaded comment ${commentIndex} id`);
        if (ids.has(id)) throw new Error(`Duplicate Spreadsheet threaded comment id ${id}.`);
        ids.add(id);
        if (commentIndex === 0) rootId = id;
        const parentId = commentIndex === 0 ? undefined : normalizeGuid(comment.parentId || rootId, `Spreadsheet threaded comment ${id} parentId`);
        const person = commentPerson(comment, thread);
        const priorPerson = people.get(person.id);
        if (priorPerson && JSON.stringify(priorPerson) !== JSON.stringify(person)) throw new Error(`Spreadsheet person ${person.id} has conflicting metadata.`);
        people.set(person.id, person);
        entries.push({ id, parentId, personId: person.id, ref, date: commentDate(comment), done: Boolean(comment.done ?? thread.resolved), text: String(comment.text ?? "") });
      }
    }
    const byId = new Map(entries.map((entry) => [entry.id, entry]));
    for (const entry of entries) {
      if (entry.parentId && !byId.has(entry.parentId)) throw new Error(`Spreadsheet threaded comment ${entry.id} references missing parent ${entry.parentId}.`);
      const visited = new Set([entry.id]);
      let parentId = entry.parentId;
      while (parentId) {
        if (visited.has(parentId)) throw new Error(`Spreadsheet threaded comment ${entry.id} has a cyclic parent chain.`);
        visited.add(parentId);
        const parent = byId.get(parentId);
        if (!parent) break;
        if (parent.ref !== entry.ref) throw new Error(`Spreadsheet threaded comment ${entry.id} and parent ${parent.id} must target the same cell.`);
        parentId = parent.parentId;
      }
    }
    return { ...part, entries };
  });
  return { parts, people: [...people.values()] };
}

export function spreadsheetThreadedCommentsXml(part) {
  const comments = part.entries.map((entry) => `<threadedComment ref="${attrEscape(entry.ref)}" dT="${attrEscape(entry.date)}" personId="${attrEscape(entry.personId)}" id="${attrEscape(entry.id)}"${entry.parentId ? ` parentId="${attrEscape(entry.parentId)}"` : ""} done="${entry.done ? 1 : 0}"><text>${attrEscape(entry.text)}</text></threadedComment>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><ThreadedComments xmlns="${XLSX_THREADED_COMMENTS_NAMESPACE}">${comments}</ThreadedComments>`;
}

export function spreadsheetPersonsXml(plan) {
  const people = plan.people.map((person) => `<person displayName="${attrEscape(person.displayName)}" id="${attrEscape(person.id)}" userId="${attrEscape(person.userId)}" providerId="${attrEscape(person.providerId)}"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><personList xmlns="${XLSX_THREADED_COMMENTS_NAMESPACE}">${people}</personList>`;
}

export function parseSpreadsheetPeople(xml = "") {
  const people = new Map();
  for (const match of String(xml || "").matchAll(/<(?:[A-Za-z_][\w.-]*:)?person\b[^>]*\/?\s*>/g)) {
    const id = localAttribute(match[0], "id")?.toUpperCase();
    if (!id || people.has(id)) continue;
    people.set(id, { id, displayName: localAttribute(match[0], "displayName") || "User", userId: localAttribute(match[0], "userId") || undefined, providerId: localAttribute(match[0], "providerId") || undefined });
  }
  return people;
}

export function parseSpreadsheetThreadedComments(xml = "") {
  const entries = [];
  for (const match of String(xml || "").matchAll(/<(?:[A-Za-z_][\w.-]*:)?threadedComment\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?threadedComment>/g)) {
    const opening = /^<(?:[A-Za-z_][\w.-]*:)?threadedComment\b[^>]*>/.exec(match[0])?.[0] || "";
    const done = String(localAttribute(opening, "done") || "0").toLowerCase();
    entries.push({
      id: localAttribute(opening, "id")?.toUpperCase(),
      parentId: localAttribute(opening, "parentId")?.toUpperCase(),
      personId: localAttribute(opening, "personId")?.toUpperCase(),
      ref: localAttribute(opening, "ref")?.toUpperCase(),
      date: localAttribute(opening, "dT") || undefined,
      done: new Set(["1", "true", "on"]).has(done),
      text: decodeXml(/<(?:[A-Za-z_][\w.-]*:)?text\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?text>/.exec(match[0])?.[1] || ""),
    });
  }
  return entries;
}

export function spreadsheetThreadRoots(entries = []) {
  const byId = new Map(entries.map((entry) => [entry.id, entry]));
  const rootId = (entry) => {
    const visited = new Set();
    let current = entry;
    while (current?.parentId && byId.has(current.parentId) && !visited.has(current.id)) {
      visited.add(current.id);
      current = byId.get(current.parentId);
    }
    return current?.id || entry.id;
  };
  const groups = new Map();
  for (const entry of entries) {
    const key = rootId(entry);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(entry);
  }
  return [...groups.values()].map((group) => group.sort((a, b) => entries.indexOf(a) - entries.indexOf(b)));
}

function issue(type, message, detail = {}) {
  return { kind: "ooxmlIssue", family: "XLSX", type, severity: "error", ...detail, message };
}

function relationshipSource(partPath) {
  if (partPath === "_rels/.rels") return "";
  const match = /^(?:(.*)\/)?_rels\/([^/]+)\.rels$/.exec(partPath);
  if (!match) return undefined;
  return match[1] ? `${match[1]}/${match[2]}` : match[2];
}

function resolveTarget(source, rawTarget) {
  const target = String(rawTarget || "").split("#")[0];
  if (target.startsWith("/")) return target.slice(1);
  const sourceDirectory = source ? path.posix.dirname(source) : "";
  return path.posix.normalize(path.posix.join(sourceDirectory === "." ? "" : sourceDirectory, target));
}

function packageXml(bytesByPath, partPath) {
  const bytes = bytesByPath.get(partPath);
  return bytes ? decoder.decode(bytes) : "";
}

function declaredContentType(contentTypes, partPath) {
  const extension = path.posix.extname(partPath).slice(1).toLowerCase();
  return contentTypes.overrides.get(partPath) || contentTypes.defaults.get(extension);
}

function relationships(bytesByPath) {
  const result = [];
  for (const [partPath] of bytesByPath) {
    if (!partPath.endsWith(".rels")) continue;
    const source = relationshipSource(partPath);
    if (source == null) continue;
    for (const match of packageXml(bytesByPath, partPath).matchAll(/<(?:[A-Za-z_][\w.-]*:)?Relationship\b[^>]*\/?\s*>/g)) {
      const attrs = attributes(match[0]);
      if (![XLSX_PERSON_RELATIONSHIP_TYPE, XLSX_THREADED_COMMENTS_RELATIONSHIP_TYPE].includes(attrs.Type)) continue;
      result.push({ source, relationshipPath: partPath, id: attrs.Id, type: attrs.Type, target: attrs.Target ? resolveTarget(source, attrs.Target) : undefined, external: String(attrs.TargetMode || "").toLowerCase() === "external" });
    }
  }
  return result;
}

function rootNamespace(xml, localName) {
  const root = rootTag(xml, localName);
  if (!root) return undefined;
  const attrs = attributes(root[0]);
  return root[1] ? attrs[`xmlns:${root[1]}`] : attrs.xmlns;
}

export function validateSpreadsheetThreadedCommentPackageSemantics({ bytesByPath, contentTypes }) {
  const issues = [];
  const rels = relationships(bytesByPath);
  const personRels = rels.filter((rel) => rel.type === XLSX_PERSON_RELATIONSHIP_TYPE);
  const threadRels = rels.filter((rel) => rel.type === XLSX_THREADED_COMMENTS_RELATIONSHIP_TYPE);
  const people = new Map();
  for (const rel of personRels) {
    if (rel.external) issues.push(issue("xlsxPersonRelationshipExternal", "XLSX person relationship must be internal.", { path: rel.relationshipPath, relationshipId: rel.id }));
    if (!rootTag(packageXml(bytesByPath, rel.source), "workbook")) issues.push(issue("xlsxPersonRelationshipSourceInvalid", "XLSX person relationship must originate from a workbook part.", { path: rel.relationshipPath, source: rel.source }));
    if (!rel.target || !bytesByPath.has(rel.target)) continue;
    if (declaredContentType(contentTypes, rel.target) !== XLSX_PERSON_CONTENT_TYPE) issues.push(issue("xlsxPersonContentTypeInvalid", `XLSX person part ${rel.target} has the wrong content type.`, { path: rel.target }));
    const xml = packageXml(bytesByPath, rel.target);
    if (rootNamespace(xml, "personList") !== XLSX_THREADED_COMMENTS_NAMESPACE) issues.push(issue("xlsxPersonRootInvalid", `XLSX person part ${rel.target} must have a personList root in the Office 2019 threaded-comments namespace.`, { path: rel.target }));
    for (const match of xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?person\b[^>]*\/?\s*>/g)) {
      const id = localAttribute(match[0], "id")?.toUpperCase();
      const displayName = localAttribute(match[0], "displayName");
      if (!id || !GUID.test(id)) issues.push(issue("xlsxPersonIdInvalid", `XLSX person id ${id || "(missing)"} must be a brace-delimited GUID.`, { path: rel.target, personId: id }));
      if (!displayName) issues.push(issue("xlsxPersonDisplayNameMissing", "XLSX person requires displayName.", { path: rel.target, personId: id }));
      if (!id) continue;
      if (people.has(id)) issues.push(issue("xlsxPersonIdDuplicate", `XLSX person id ${id} is duplicated.`, { path: rel.target, personId: id }));
      people.set(id, { id, displayName: displayName || "User", userId: localAttribute(match[0], "userId") || undefined, providerId: localAttribute(match[0], "providerId") || undefined });
    }
  }
  const personCountBySource = new Map();
  for (const rel of personRels) personCountBySource.set(rel.source, (personCountBySource.get(rel.source) || 0) + 1);
  for (const [source, count] of personCountBySource) if (count > 1) issues.push(issue("xlsxPersonRelationshipMultiplicity", `XLSX workbook ${source} has ${count} person relationships; at most one is allowed.`, { source, count }));

  const globalCommentIds = new Set();
  const threadCountBySource = new Map();
  for (const rel of threadRels) {
    threadCountBySource.set(rel.source, (threadCountBySource.get(rel.source) || 0) + 1);
    if (rel.external) issues.push(issue("xlsxThreadedCommentRelationshipExternal", "XLSX threadedComment relationship must be internal.", { path: rel.relationshipPath, relationshipId: rel.id }));
    if (!rootTag(packageXml(bytesByPath, rel.source), "worksheet")) issues.push(issue("xlsxThreadedCommentRelationshipSourceInvalid", "XLSX threadedComment relationship must originate from a worksheet part.", { path: rel.relationshipPath, source: rel.source }));
    if (!rel.target || !bytesByPath.has(rel.target)) continue;
    if (declaredContentType(contentTypes, rel.target) !== XLSX_THREADED_COMMENTS_CONTENT_TYPE) issues.push(issue("xlsxThreadedCommentContentTypeInvalid", `XLSX threaded-comments part ${rel.target} has the wrong content type.`, { path: rel.target }));
    const xml = packageXml(bytesByPath, rel.target);
    if (rootNamespace(xml, "ThreadedComments") !== XLSX_THREADED_COMMENTS_NAMESPACE) issues.push(issue("xlsxThreadedCommentsRootInvalid", `XLSX threaded-comments part ${rel.target} must have a ThreadedComments root in the Office 2019 namespace.`, { path: rel.target }));
    const entries = parseSpreadsheetThreadedComments(xml);
    const byId = new Map();
    for (const entry of entries) {
      if (!entry.id || !GUID.test(entry.id)) issues.push(issue("xlsxThreadedCommentIdInvalid", `XLSX threaded comment id ${entry.id || "(missing)"} must be a brace-delimited GUID.`, { path: rel.target, commentId: entry.id }));
      else {
        if (byId.has(entry.id) || globalCommentIds.has(entry.id)) issues.push(issue("xlsxThreadedCommentIdDuplicate", `XLSX threaded comment id ${entry.id} is duplicated.`, { path: rel.target, commentId: entry.id }));
        byId.set(entry.id, entry);
        globalCommentIds.add(entry.id);
      }
      if (!entry.personId || !GUID.test(entry.personId)) issues.push(issue("xlsxThreadedCommentPersonIdInvalid", `XLSX threaded comment personId ${entry.personId || "(missing)"} must be a brace-delimited GUID.`, { path: rel.target, personId: entry.personId }));
      else if (!people.has(entry.personId)) issues.push(issue("xlsxThreadedCommentPersonNotFound", `XLSX threaded comment references missing person ${entry.personId}.`, { path: rel.target, personId: entry.personId }));
      if (entry.parentId && !GUID.test(entry.parentId)) issues.push(issue("xlsxThreadedCommentParentIdInvalid", `XLSX threaded comment parentId ${entry.parentId} must be a brace-delimited GUID.`, { path: rel.target, parentId: entry.parentId }));
      if (entry.date && Number.isNaN(Date.parse(entry.date))) issues.push(issue("xlsxThreadedCommentDateInvalid", `XLSX threaded comment ${entry.id} has invalid dT ${entry.date}.`, { path: rel.target, commentId: entry.id, date: entry.date }));
      if (!entry.parentId && !entry.ref) issues.push(issue("xlsxThreadedCommentRefMissing", `XLSX root threaded comment ${entry.id || "(missing)"} requires ref.`, { path: rel.target, commentId: entry.id }));
    }
    for (const entry of entries) {
      if (entry.parentId && !byId.has(entry.parentId)) issues.push(issue("xlsxThreadedCommentParentNotFound", `XLSX threaded comment ${entry.id} references missing parent ${entry.parentId}.`, { path: rel.target, commentId: entry.id, parentId: entry.parentId }));
      const visited = new Set([entry.id]);
      let parentId = entry.parentId;
      while (parentId && byId.has(parentId)) {
        if (visited.has(parentId)) { issues.push(issue("xlsxThreadedCommentParentCycle", `XLSX threaded-comments part ${rel.target} contains a parent cycle.`, { path: rel.target, commentId: entry.id, parentId })); break; }
        visited.add(parentId);
        const parent = byId.get(parentId);
        if (entry.ref && parent.ref && entry.ref !== parent.ref) issues.push(issue("xlsxThreadedCommentParentRefMismatch", `XLSX threaded comment ${entry.id} and parent ${parent.id} target different cells.`, { path: rel.target, commentId: entry.id, parentId: parent.id }));
        parentId = parent.parentId;
      }
    }
  }
  for (const [source, count] of threadCountBySource) if (count > 1) issues.push(issue("xlsxThreadedCommentRelationshipMultiplicity", `XLSX worksheet ${source} has ${count} threadedComment relationships; at most one is allowed.`, { source, count }));
  if (threadRels.length && !personRels.length) issues.push(issue("xlsxPersonPartMissing", "XLSX threaded comments require a workbook person relationship."));

  const personTargets = new Set(personRels.filter((rel) => !rel.external).map((rel) => rel.target));
  const threadTargets = new Set(threadRels.filter((rel) => !rel.external).map((rel) => rel.target));
  for (const partPath of bytesByPath.keys()) {
    const type = declaredContentType(contentTypes, partPath);
    if (type === XLSX_PERSON_CONTENT_TYPE && !personTargets.has(partPath)) issues.push(issue("xlsxPersonPartOrphaned", `XLSX person part ${partPath} is not targeted by a workbook relationship.`, { path: partPath }));
    if (type === XLSX_THREADED_COMMENTS_CONTENT_TYPE && !threadTargets.has(partPath)) issues.push(issue("xlsxThreadedCommentsPartOrphaned", `XLSX threaded-comments part ${partPath} is not targeted by a worksheet relationship.`, { path: partPath }));
  }
  return issues;
}
