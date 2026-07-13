function padded(value, width = 2) {
  return String(value).padStart(width, "0");
}

const MAX_DATE_SERIAL = { "1900": 2_958_465, "1904": 2_957_003 };

function dateAtUtc(year, month, day) {
  const value = new Date(0);
  value.setUTCFullYear(year, month - 1, day);
  value.setUTCHours(0, 0, 0, 0);
  return value;
}

function validCalendarDate(year, month, day) {
  const calendar = dateAtUtc(year, month, day);
  return calendar.getUTCFullYear() === year && calendar.getUTCMonth() === month - 1 && calendar.getUTCDate() === day;
}

export function pivotDateSerial(yearValue, monthValue, dayValue = 1, dateSystem = "1900") {
  let year = Math.trunc(Number(yearValue));
  const month = Math.trunc(Number(monthValue));
  const day = Math.trunc(Number(dayValue));
  if (![year, month, day].every(Number.isFinite) || year < 0 || year >= 10_000) return undefined;
  if (year <= 1_899) year += 1_900;
  const normalized = dateAtUtc(year, month, 1);
  if (normalized.getUTCFullYear() < 1_900 || normalized.getUTCFullYear() >= 10_000) return undefined;
  const epoch = dateSystem === "1904" ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31);
  let serial = Math.floor((normalized.getTime() - epoch) / 86_400_000) + day - 1;
  if (dateSystem !== "1904" && normalized.getTime() >= Date.UTC(1900, 2, 1)) serial += 1;
  return serial < 0 || serial > MAX_DATE_SERIAL[dateSystem === "1904" ? "1904" : "1900"] ? undefined : serial;
}

export function pivotMaxDateSerial(dateSystem = "1900") {
  return MAX_DATE_SERIAL[dateSystem === "1904" ? "1904" : "1900"];
}

export function pivotFormulaDateParts(value, dateSystem = "1900") {
  const serial = Math.floor(Number(value));
  if (!Number.isFinite(serial) || serial < 0 || serial > pivotMaxDateSerial(dateSystem)) return undefined;
  if (dateSystem !== "1904" && serial === 0) return { year: 1900, month: 1, day: 0 };
  return pivotDateParts(serial, dateSystem);
}

function daysInMonth(year, month, dateSystem) {
  if (dateSystem !== "1904" && year === 1900 && month === 2) return 29;
  return dateAtUtc(year, month + 1, 0).getUTCDate();
}

export function pivotShiftDateMonths(serialValue, monthsValue, endOfMonth = false, dateSystem = "1900") {
  const parts = pivotFormulaDateParts(serialValue, dateSystem);
  const months = Math.trunc(Number(monthsValue));
  if (!parts || !Number.isFinite(months)) return undefined;
  const first = dateAtUtc(parts.year, parts.month + months, 1);
  const year = first.getUTCFullYear();
  const month = first.getUTCMonth() + 1;
  if (year < 1900 || year >= 10_000) return undefined;
  const day = endOfMonth ? daysInMonth(year, month, dateSystem) : Math.min(Math.max(1, parts.day), daysInMonth(year, month, dateSystem));
  return pivotDateSerial(year, month, day, dateSystem);
}

export function pivotWeekdayIndex(serialValue, dateSystem = "1900") {
  const serial = Math.floor(Number(serialValue));
  if (!pivotFormulaDateParts(serial, dateSystem)) return undefined;
  if (dateSystem === "1904") return ((serial + 5) % 7 + 7) % 7;
  const adjusted = serial > 60 ? serial - 1 : serial;
  return ((adjusted % 7) + 7) % 7;
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

export function normalizePivotDateTime(value, label) {
  const result = pivotDateTimeKey(value);
  if (!result || typeof value === "number") throw new TypeError(`${label} must be an ISO date-time or Date.`);
  return result;
}
