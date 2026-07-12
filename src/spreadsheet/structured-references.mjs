// Pure Excel structured-reference syntax parsing. Workbook/table resolution stays in the model layer.

const TABLE_NAME = /^[A-Za-z_\\][A-Za-z0-9_.\\]*/;
const ESCAPED_COLUMN_CHARACTER = new Set(["'", "[", "]", "#", "@"]);

function matchingBracketEnd(text, start) {
  let depth = 0;
  for (let index = start; index < text.length; index += 1) {
    const character = text[index];
    if (character === "'" && ESCAPED_COLUMN_CHARACTER.has(text[index + 1])) {
      index += 1;
      continue;
    }
    if (character === "[") depth += 1;
    else if (character === "]" && --depth === 0) return index + 1;
  }
  return -1;
}

function decodeColumnToken(raw = "") {
  let value = "";
  for (let index = 0; index < raw.length; index += 1) {
    if (raw[index] === "'" && ESCAPED_COLUMN_CHARACTER.has(raw[index + 1])) index += 1;
    value += raw[index];
  }
  return value.trim();
}

function innerTokens(body = "") {
  const tokens = [];
  let separator;
  for (let index = 0; index < body.length;) {
    const character = body[index];
    if (/\s/.test(character) || character === ",") {
      if (character === ",") separator = ",";
      index += 1;
      continue;
    }
    if (character === ":") {
      separator = ":";
      index += 1;
      continue;
    }
    if (character === "@" && body[index + 1] === "[") {
      tokens.push({ raw: "@", value: "@", separator });
      separator = undefined;
      index += 1;
      continue;
    }
    if (character === "[") {
      const end = matchingBracketEnd(body, index);
      if (end < 0) return undefined;
      const raw = body.slice(index + 1, end - 1).trim();
      tokens.push({ raw, value: decodeColumnToken(raw), separator });
      separator = undefined;
      index = end;
      continue;
    }
    let end = index;
    while (end < body.length) {
      if (body[end] === "'" && ESCAPED_COLUMN_CHARACTER.has(body[end + 1])) {
        end += 2;
        continue;
      }
      if (/[,:\[\]]/.test(body[end])) break;
      end += 1;
    }
    const raw = body.slice(index, end).trim();
    if (raw) tokens.push({ raw, value: decodeColumnToken(raw), separator });
    separator = undefined;
    index = end;
  }
  return tokens;
}

export function parseStructuredReference(reference = "") {
  const text = String(reference || "").trim();
  if (!text) return undefined;
  let tableName;
  let bracketStart = 0;
  let explicitAt = false;
  if (text.startsWith("@[")) {
    explicitAt = true;
    bracketStart = 1;
  } else if (!text.startsWith("[")) {
    const table = TABLE_NAME.exec(text);
    if (!table || text[table[0].length] !== "[") return undefined;
    tableName = table[0];
    bracketStart = table[0].length;
  }
  const end = matchingBracketEnd(text, bracketStart);
  if (end !== text.length) return undefined;
  const body = text.slice(bracketStart + 1, -1).trim();
  const tokens = innerTokens(body);
  if (!tokens?.length) return undefined;
  const sectionTokens = [];
  const columnTokens = [];
  let currentRow = explicitAt;
  for (const token of tokens) {
    const escapedSpecial = token.raw.startsWith("'") && ESCAPED_COLUMN_CHARACTER.has(token.raw[1]);
    if (!escapedSpecial && /^#This Row$/i.test(token.value)) {
      currentRow = true;
      sectionTokens.push("#This Row");
      continue;
    }
    if (!escapedSpecial && token.value === "@") {
      currentRow = true;
      continue;
    }
    if (!escapedSpecial && token.value.startsWith("@")) {
      currentRow = true;
      const value = token.value.slice(1).trim();
      if (value) columnTokens.push({ ...token, value });
      continue;
    }
    if (!escapedSpecial && token.value.startsWith("#")) sectionTokens.push(token.value);
    else columnTokens.push(token);
  }
  const columnSelectors = [];
  for (let index = 0; index < columnTokens.length; index += 1) {
    const token = columnTokens[index];
    const next = columnTokens[index + 1];
    if (next?.separator === ":") {
      columnSelectors.push({ start: token.value, end: next.value });
      index += 1;
    } else columnSelectors.push({ name: token.value });
  }
  return {
    text,
    tableName,
    qualified: Boolean(tableName),
    currentRow,
    sectionTokens,
    columnSelectors,
    tokens: tokens.map((token) => token.value),
  };
}

export function scanStructuredReferences(formula = "") {
  const text = String(formula || "");
  const references = [];
  let quoted = false;
  for (let index = 0; index < text.length;) {
    if (text[index] === '"') {
      if (quoted && text[index + 1] === '"') index += 2;
      else {
        quoted = !quoted;
        index += 1;
      }
      continue;
    }
    if (quoted) {
      index += 1;
      continue;
    }
    const start = index;
    let bracketStart = -1;
    if (text[index] === "[") bracketStart = index;
    else if (text[index] === "@" && text[index + 1] === "[") bracketStart = index + 1;
    else if ((index === 0 || !/[A-Za-z0-9_.\\]/.test(text[index - 1])) && /[A-Za-z_\\]/.test(text[index])) {
      const table = TABLE_NAME.exec(text.slice(index));
      if (table && text[index + table[0].length] === "[") bracketStart = index + table[0].length;
    }
    if (bracketStart < 0) {
      index += 1;
      continue;
    }
    const end = matchingBracketEnd(text, bracketStart);
    if (end < 0) {
      index += 1;
      continue;
    }
    const reference = text.slice(start, end);
    const parsed = parseStructuredReference(reference);
    if (parsed) references.push({ ...parsed, start, end });
    index = end;
  }
  return references;
}

export function scanStructuredReferenceIntersections(formula = "") {
  const text = String(formula || "");
  const references = scanStructuredReferences(text);
  const intersections = [];
  for (let index = 0; index < references.length;) {
    const group = [references[index]];
    let cursor = index + 1;
    while (cursor < references.length && /^\s+$/.test(text.slice(group.at(-1).end, references[cursor].start))) {
      group.push(references[cursor]);
      cursor += 1;
    }
    if (group.length > 1) {
      intersections.push({
        text: text.slice(group[0].start, group.at(-1).end),
        start: group[0].start,
        end: group.at(-1).end,
        references: group,
      });
    }
    index = cursor;
  }
  return intersections;
}

export function splitReferenceIntersectionOperands(reference = "") {
  const text = String(reference || "").trim();
  const operands = [];
  let current = "";
  let bracketDepth = 0;
  let doubleQuoted = false;
  let singleQuoted = false;
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    if (character === '"' && !singleQuoted) {
      current += character;
      if (doubleQuoted && text[index + 1] === '"') {
        current += '"';
        index += 1;
      } else doubleQuoted = !doubleQuoted;
      continue;
    }
    if (character === "'" && !doubleQuoted && bracketDepth === 0) {
      current += character;
      if (singleQuoted && text[index + 1] === "'") {
        current += "'";
        index += 1;
      } else singleQuoted = !singleQuoted;
      continue;
    }
    if (!doubleQuoted && !singleQuoted) {
      if (character === "[") bracketDepth += 1;
      else if (character === "]") bracketDepth = Math.max(0, bracketDepth - 1);
      if (/\s/.test(character) && bracketDepth === 0) {
        if (current.trim()) operands.push(current.trim());
        current = "";
        continue;
      }
    }
    current += character;
  }
  if (current.trim()) operands.push(current.trim());
  return operands.length > 1 ? operands : undefined;
}
