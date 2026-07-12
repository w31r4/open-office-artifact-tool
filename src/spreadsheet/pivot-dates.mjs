function padded(value, width = 2) {
  return String(value).padStart(width, "0");
}

function validCalendarDate(year, month, day) {
  const calendar = new Date(0);
  calendar.setUTCFullYear(year, month - 1, day);
  calendar.setUTCHours(0, 0, 0, 0);
  return calendar.getUTCFullYear() === year && calendar.getUTCMonth() === month - 1 && calendar.getUTCDate() === day;
}

function textDateParts(value) {
  const text = String(value ?? "").trim();
  const match = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2})(?::(\d{2}))?(?::(\d{2})(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})?)?$/.exec(text);
  if (!match) return undefined;
  const [year, month, day, hour = 0, minute = 0, second = 0] = match.slice(1).map((item) => Number(item || 0));
  if (!validCalendarDate(year, month, day) || hour > 23 || minute > 59 || second > 59) return undefined;
  if (/[Zz]|[+-]\d{2}:?\d{2}$/.test(text)) {
    const instant = new Date(text);
    if (Number.isNaN(instant.valueOf())) return undefined;
    return dateParts(instant);
  }
  return { year, month, day, hour, minute, second };
}

function serialDateParts(value, dateSystem) {
  const wholeDays = Math.floor(value);
  const fraction = value - wholeDays;
  const timeMilliseconds = Math.round(fraction * 86_400_000);
  if (dateSystem !== "1904" && wholeDays === 60) {
    const time = new Date(timeMilliseconds);
    return { year: 1900, month: 2, day: 29, hour: time.getUTCHours(), minute: time.getUTCMinutes(), second: time.getUTCSeconds() };
  }
  const epoch = Date.UTC(dateSystem === "1904" ? 1904 : 1899, dateSystem === "1904" ? 0 : 11, dateSystem === "1904" ? 1 : 31);
  const adjustedDays = dateSystem === "1904" ? wholeDays : wholeDays - (wholeDays > 60 ? 1 : 0);
  return dateParts(new Date(epoch + adjustedDays * 86_400_000 + timeMilliseconds));
}

function dateParts(value) {
  if (Number.isNaN(value.valueOf())) return undefined;
  return {
    year: value.getUTCFullYear(),
    month: value.getUTCMonth() + 1,
    day: value.getUTCDate(),
    hour: value.getUTCHours(),
    minute: value.getUTCMinutes(),
    second: value.getUTCSeconds(),
  };
}

export function pivotDateParts(value, dateSystem = "1900") {
  if (value instanceof Date) return dateParts(value);
  if (typeof value === "number" && Number.isFinite(value)) return serialDateParts(value, dateSystem);
  return textDateParts(value);
}

export function pivotDateKey(value, dateSystem = "1900") {
  const parts = pivotDateParts(value, dateSystem);
  return parts ? `${padded(parts.year, 4)}-${padded(parts.month)}-${padded(parts.day)}` : undefined;
}

export function pivotDateTimeKey(value, dateSystem = "1900") {
  const parts = pivotDateParts(value, dateSystem);
  return parts ? `${padded(parts.year, 4)}-${padded(parts.month)}-${padded(parts.day)}T${padded(parts.hour)}:${padded(parts.minute)}:${padded(parts.second)}` : undefined;
}

export function normalizePivotDate(value, label) {
  const result = pivotDateKey(value);
  if (!result || typeof value === "number") throw new TypeError(`${label} must be an ISO date or Date.`);
  return result;
}
