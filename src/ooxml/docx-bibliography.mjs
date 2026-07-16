import path from "node:path";
import { attributes, decodeXml, rootTag } from "./source-reference-xml.mjs";

const SOURCE_TYPES = new Set(["ArticleInAPeriodical", "Book", "BookSection", "JournalArticle", "ConferenceProceedings", "Report", "SoundRecording", "Performance", "Art", "DocumentFromInternetSite", "InternetSite", "Film", "Interview", "Patent", "ElectronicSource", "Case", "Misc"]);
const FIELD_TAGS = {
  title: "Title", year: "Year", city: "City", stateProvince: "StateProvince", countryRegion: "CountryRegion", publisher: "Publisher",
  bookTitle: "BookTitle", journalName: "JournalName", periodicalTitle: "PeriodicalTitle", publicationTitle: "PublicationTitle", internetSiteTitle: "InternetSiteTitle",
  conferenceName: "ConferenceName", institution: "Institution", department: "Department", volume: "Volume", issue: "Issue", pages: "Pages", edition: "Edition",
  numberVolumes: "NumberVolumes", chapterNumber: "ChapterNumber", standardNumber: "StandardNumber", shortTitle: "ShortTitle", comments: "Comments", medium: "Medium",
  month: "Month", day: "Day", yearAccessed: "YearAccessed", monthAccessed: "MonthAccessed", dayAccessed: "DayAccessed", url: "URL", guid: "Guid", lcid: "LCID",
  reporter: "Reporter", caseNumber: "CaseNumber", abbreviatedCaseNumber: "AbbreviatedCaseNumber", court: "Court", patentNumber: "PatentNumber", patentType: "Type",
  broadcaster: "Broadcaster", broadcastTitle: "BroadcastTitle", station: "Station", theater: "Theater", productionCompany: "ProductionCompany", distributor: "Distributor",
  recordingNumber: "RecordingNumber", albumTitle: "AlbumTitle", thesisType: "ThesisType", version: "Version", referenceOrder: "RefOrder",
};
const FIELD_BY_TAG = new Map(Object.entries(FIELD_TAGS).map(([name, tag]) => [tag, name]));
const decoder = new TextDecoder();

function localAttribute(tag, localName) {
  return Object.entries(attributes(tag)).find(([name]) => name === localName || name.endsWith(`:${localName}`))?.[1];
}

function string255(value, label, required = false) {
  const text = String(value ?? "").trim();
  if (required && !text) throw new TypeError(`${label} must be non-empty.`);
  if (text.length > 255) throw new RangeError(`${label} must contain at most 255 characters.`);
  return text;
}

function stableGuid(tag) {
  const hash = (seed) => {
    let value = seed >>> 0;
    for (const character of String(tag)) value = Math.imul(value ^ character.codePointAt(0), 16777619) >>> 0;
    return value.toString(16).padStart(8, "0").toUpperCase();
  };
  const hex = `${hash(2166136261)}${hash(2246822519)}${hash(3266489917)}${hash(668265263)}`;
  return `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-4${hex.slice(13, 16)}-8${hex.slice(17, 20)}-${hex.slice(20, 32)}}`;
}

function normalizeAuthor(author, index, tag) {
  if (typeof author === "string") {
    const parts = author.trim().split(/\s+/).filter(Boolean);
    return { first: string255(parts.slice(0, -1).join(" "), `DOCX bibliography source ${tag} author ${index} first name`), last: string255(parts.at(-1) || "", `DOCX bibliography source ${tag} author ${index} last name`) };
  }
  return {
    last: string255(author?.last ?? author?.family, `DOCX bibliography source ${tag} author ${index} last name`),
    first: string255(author?.first ?? author?.given, `DOCX bibliography source ${tag} author ${index} first name`),
    middle: string255(author?.middle, `DOCX bibliography source ${tag} author ${index} middle name`),
  };
}

export function normalizeDocxBibliographySource(config = {}, index = 0) {
  const tag = string255(config.tag ?? config.bibliographyTag ?? config.id, `DOCX bibliography source ${index} tag`, true);
  const sourceType = string255(config.sourceType ?? config.type ?? (config.url ? "InternetSite" : "Misc"), `DOCX bibliography source ${tag} type`, true);
  if (!SOURCE_TYPES.has(sourceType)) throw new TypeError(`DOCX bibliography source ${tag} has unsupported SourceType ${sourceType}.`);
  const fields = {};
  for (const [name] of Object.entries(FIELD_TAGS)) {
    const value = config[name] ?? config.fields?.[name];
    if (value !== undefined && value !== null && String(value).length) fields[name] = string255(value, `DOCX bibliography source ${tag} ${name}`);
  }
  if (!fields.guid) fields.guid = stableGuid(tag);
  const rawAuthors = config.authors ?? config.author ?? [];
  const authors = (Array.isArray(rawAuthors) ? rawAuthors : [rawAuthors]).map((author, authorIndex) => normalizeAuthor(author, authorIndex, tag)).filter((author) => author.first || author.middle || author.last);
  const corporateAuthor = string255(config.corporateAuthor ?? config.corporate, `DOCX bibliography source ${tag} corporate author`);
  if (authors.length && corporateAuthor) throw new TypeError(`DOCX bibliography source ${tag} cannot combine personal authors with a corporate author.`);
  return { id: config.id || `bibliography/${tag}`, kind: "bibliographySource", tag, sourceType, authors, corporateAuthor: corporateAuthor || undefined, ...fields };
}

function elementText(xml, localName) {
  const escaped = String(localName).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return decodeXml(new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${escaped}\\b[^>]*>([\\s\\S]*?)<\\/(?:[A-Za-z_][\\w.-]*:)?${escaped}>`).exec(xml)?.[1] || "").trim();
}

export function parseDocxBibliography(xml = "") {
  const source = String(xml || "");
  if (!rootTag(source, "Sources")) return { entries: [], byTag: new Map() };
  const opening = /<(?:[A-Za-z_][\w.-]*:)?Sources\b[^>]*>/.exec(source)?.[0] || "";
  const entries = [...source.matchAll(/<(?:[A-Za-z_][\w.-]*:)?Source\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?Source>/g)].map((match, index) => {
    const body = match[1];
    const config = { tag: elementText(body, "Tag"), sourceType: elementText(body, "SourceType") };
    for (const [tag, name] of FIELD_BY_TAG) {
      const value = elementText(body, tag);
      if (value) config[name] = value;
    }
    config.authors = [...body.matchAll(/<(?:[A-Za-z_][\w.-]*:)?Person\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?Person>/g)].map((person) => ({ last: elementText(person[1], "Last"), first: elementText(person[1], "First"), middle: elementText(person[1], "Middle") })).filter((author) => author.first || author.middle || author.last);
    config.corporateAuthor = elementText(body, "Corporate") || undefined;
    return normalizeDocxBibliographySource(config, index);
  });
  const byTag = new Map();
  for (const entry of entries) {
    if (byTag.has(entry.tag)) throw new Error(`Duplicate DOCX bibliography source tag ${entry.tag}.`);
    byTag.set(entry.tag, entry);
  }
  return {
    entries,
    byTag,
    selectedStyle: string255(decodeXml(localAttribute(opening, "SelectedStyle") || ""), "DOCX bibliography SelectedStyle"),
    styleName: string255(decodeXml(localAttribute(opening, "StyleName") || ""), "DOCX bibliography StyleName"),
    uri: string255(decodeXml(localAttribute(opening, "URI") || ""), "DOCX bibliography URI"),
  };
}

export function parseDocxCitationInstruction(instruction = "") {
  const match = /^\s*CITATION\s+(?:"([^"]+)"|'([^']+)'|([^\s\\]+))/i.exec(decodeXml(String(instruction || "")));
  return match ? String(match[1] || match[2] || match[3] || "").trim() : undefined;
}

function issue(type, message, detail = {}) {
  return { kind: "ooxmlIssue", family: "DOCX", type, severity: "error", ...detail, message };
}

function resolveTarget(sourcePath, target) {
  return path.posix.normalize(path.posix.join(path.posix.dirname(sourcePath), String(target || "").replace(/^\//, ""))).replace(/^\.\//, "");
}

export function validateDocxBibliographyPackageSemantics({ bytesByPath }) {
  const issues = [];
  const bibliographyParts = [];
  for (const [partPath, bytes] of bytesByPath) {
    const xml = decoder.decode(bytes);
    if (!rootTag(xml, "Sources")) continue;
    try {
      bibliographyParts.push({ partPath, plan: parseDocxBibliography(xml) });
    } catch (error) {
      issues.push(issue("docxBibliographyInvalid", error.message, { path: partPath }));
    }
  }
  if (bibliographyParts.length > 1) issues.push(issue("docxBibliographyMultipleParts", "DOCX package contains more than one bibliography Sources part.", { paths: bibliographyParts.map((part) => part.partPath) }));
  const relsPath = "word/_rels/document.xml.rels";
  const relsXml = decoder.decode(bytesByPath.get(relsPath) || new Uint8Array());
  const linkedCustomXml = new Set([...relsXml.matchAll(/<Relationship\b[^>]*\/?\s*>/g)].flatMap((match) => {
    const attrs = attributes(match[0]);
    const type = attrs.Type || attrs.type || "";
    const mode = String(attrs.TargetMode || attrs.targetMode || "").toLowerCase();
    return type.endsWith("/customXml") && mode !== "external" ? [resolveTarget("word/document.xml", attrs.Target || attrs.target)] : [];
  }));
  for (const part of bibliographyParts) if (!linkedCustomXml.has(part.partPath)) issues.push(issue("docxBibliographyRelationshipMissing", `DOCX bibliography part ${part.partPath} is not linked from word/document.xml as customXml.`, { path: part.partPath }));
  const tags = bibliographyParts[0]?.plan.byTag || new Map();
  const documentXml = decoder.decode(bytesByPath.get("word/document.xml") || new Uint8Array());
  const instructions = [
    ...[...documentXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?fldSimple\b[^>]*>/g)].map((match) => localAttribute(match[0], "instr") || ""),
    ...[...documentXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?instrText\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?instrText>/g)].map((match) => decodeXml(match[1])),
  ];
  for (const instruction of instructions) {
    const tag = parseDocxCitationInstruction(instruction);
    if (tag && !tags.has(tag)) issues.push(issue("docxCitationSourceMissing", `DOCX CITATION field references missing bibliography source ${tag}.`, { path: "word/document.xml", tag }));
  }
  return issues;
}
