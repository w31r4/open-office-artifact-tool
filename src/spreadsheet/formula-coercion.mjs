const ENGLISH_MONTHS = new Map([
  ["jan", 1], ["january", 1], ["feb", 2], ["february", 2], ["mar", 3], ["march", 3],
  ["apr", 4], ["april", 4], ["may", 5], ["jun", 6], ["june", 6], ["jul", 7], ["july", 7],
  ["aug", 8], ["august", 8], ["sep", 9], ["sept", 9], ["september", 9], ["oct", 10],
  ["october", 10], ["nov", 11], ["november", 11], ["dec", 12], ["december", 12],
]);

export function parseFormulaNumberText(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : undefined;
  if (typeof value !== "string") return undefined;
  let text = value.trim();
  if (!text) return undefined;
  let negative = false;
  if (/^\(.*\)$/.test(text)) { negative = true; text = text.slice(1, -1).trim(); }
  if (negative && /^[+-]/.test(text)) return undefined;
  const percent = text.endsWith("%");
  if (percent) text = text.slice(0, -1).trim();
  if (!/^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(text)) return undefined;
  const number = Number(text.replaceAll(",", ""));
  if (!Number.isFinite(number)) return undefined;
  return (negative ? -number : number) / (percent ? 100 : 1);
}

export function parseFormulaTimeText(value) {
  if (typeof value !== "string") return undefined;
  const text = value.trim();
  const match = /(?:^|[T\s])(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?$/.exec(text);
  if (!match) return undefined;
  let hour = Number(match[1]);
  const minute = Number(match[2]);
  const second = Number(match[3] || 0);
  const meridiem = match[4]?.toUpperCase();
  if (minute > 59 || second > 59) return undefined;
  if (meridiem) {
    if (hour < 1 || hour > 12) return undefined;
    if (hour === 12) hour = 0;
    if (meridiem === "PM") hour += 12;
  } else if (hour > 23) return undefined;
  return { hour, minute, second, serial: (hour * 3600 + minute * 60 + second) / 86_400, dateText: text.slice(0, match.index).trim() || undefined };
}

export function parseFormulaDateText(value) {
  if (typeof value !== "string") return undefined;
  let text = value.trim();
  if (!text) return undefined;
  const time = parseFormulaTimeText(text);
  if (time) text = text.replace(/(?:^|[T\s])\d{1,2}:\d{2}(?::\d{2})?\s*(?:[AaPp][Mm])?$/, "").trim();
  let year, month, day;
  let match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
  if (match) [year, month, day] = match.slice(1).map(Number);
  if (!match) {
    match = /^(\d{1,2})[-\s]([A-Za-z]+)[-\s](\d{4})$/.exec(text);
    if (match) { day = Number(match[1]); month = ENGLISH_MONTHS.get(match[2].toLowerCase()); year = Number(match[3]); }
  }
  if (!match) {
    match = /^([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})$/.exec(text);
    if (match) { month = ENGLISH_MONTHS.get(match[1].toLowerCase()); day = Number(match[2]); year = Number(match[3]); }
  }
  if (!match || !month || year < 0 || year > 9999 || day < 1 || day > 31) return undefined;
  return { year, month, day, hasTime: Boolean(time) };
}

export function formulaTimeSerial(hourValue, minuteValue, secondValue) {
  const parts = [hourValue, minuteValue, secondValue].map(Number);
  if (parts.some((value) => !Number.isFinite(value) || value < 0 || value > 32767)) return undefined;
  const [hour, minute, second] = parts.map(Math.trunc);
  return ((hour * 3600 + minute * 60 + second) % 86_400) / 86_400;
}

export function formulaTimeParts(value) {
  if (typeof value === "string") return parseFormulaTimeText(value);
  const number = Number(value);
  if (!Number.isFinite(number) || number < 0) return undefined;
  const seconds = Math.round((number - Math.floor(number)) * 86_400) % 86_400;
  return { hour: Math.floor(seconds / 3600), minute: Math.floor(seconds % 3600 / 60), second: seconds % 60, serial: seconds / 86_400 };
}
