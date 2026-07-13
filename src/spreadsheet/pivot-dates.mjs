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

function pivotWeekendDays(value = 1, allowAllWeekend = false) {
  if (typeof value === "string") {
    const weekend = value;
    if (/^[01]{7}$/.test(weekend)) {
      if (weekend === "1111111" && !allowAllWeekend) return { error: "#VALUE!", weekends: new Set() };
      const weekends = new Set();
      for (let index = 0; index < 7; index += 1) if (weekend[index] === "1") weekends.add((index + 1) % 7);
      return { weekends };
    }
    return { error: "#VALUE!", weekends: new Set() };
  }
  const weekendNumber = Number(value);
  if (!Number.isInteger(weekendNumber)) return { error: "#NUM!", weekends: new Set() };
  if (weekendNumber >= 1 && weekendNumber <= 7) {
    const first = weekendNumber === 1 ? 6 : weekendNumber - 2;
    return { weekends: new Set([first, (first + 1) % 7]) };
  }
  if (weekendNumber >= 11 && weekendNumber <= 17) return { weekends: new Set([weekendNumber - 11]) };
  return { error: "#NUM!", weekends: new Set() };
}

function pivotHolidaySet(values, dateSystem) {
  const holidays = new Set();
  for (const value of values) {
    if (value == null) continue;
    const serial = Math.floor(Number(value));
    if (!pivotFormulaDateParts(serial, dateSystem)) return { error: "#NUM!", holidays: new Set() };
    holidays.add(serial);
  }
  return { holidays };
}

function pivotBusinessDaysSegment(low, high, dateSystem, weekends) {
  if (low > high) return 0;
  const total = high - low + 1;
  const fullWeeks = Math.floor(total / 7);
  let count = fullWeeks * (7 - weekends.size);
  for (let serial = low + fullWeeks * 7; serial <= high; serial += 1) {
    if (!weekends.has(pivotWeekdayIndex(serial, dateSystem))) count += 1;
  }
  return count;
}

function pivotBusinessDaysBetween(low, high, holidays, dateSystem, weekends) {
  let count = dateSystem !== "1904" && low <= 60 && high >= 61
    ? pivotBusinessDaysSegment(low, 60, dateSystem, weekends) + pivotBusinessDaysSegment(61, high, dateSystem, weekends)
    : pivotBusinessDaysSegment(low, high, dateSystem, weekends);
  for (const holiday of holidays) {
    if (holiday >= low && holiday <= high && !weekends.has(pivotWeekdayIndex(holiday, dateSystem))) count -= 1;
  }
  return count;
}

export function pivotNetworkDays(startValue, endValue, options = {}) {
  const dateSystem = options.dateSystem === "1904" ? "1904" : "1900";
  const start = Math.floor(Number(startValue));
  const end = Math.floor(Number(endValue));
  if (!pivotFormulaDateParts(start, dateSystem) || !pivotFormulaDateParts(end, dateSystem)) return "#NUM!";
  const holidayResult = pivotHolidaySet(options.holidays || [], dateSystem);
  if (holidayResult.error) return holidayResult.error;
  const weekendResult = pivotWeekendDays(options.weekend ?? 1, options.allowAllWeekend === true);
  if (weekendResult.error) return weekendResult.error;
  const direction = start <= end ? 1 : -1;
  const count = pivotBusinessDaysBetween(Math.min(start, end), Math.max(start, end), holidayResult.holidays, dateSystem, weekendResult.weekends);
  return count === 0 ? 0 : count * direction;
}

export function pivotWorkday(startValue, daysValue, options = {}) {
  const dateSystem = options.dateSystem === "1904" ? "1904" : "1900";
  const start = Math.floor(Number(startValue));
  const days = Math.trunc(Number(daysValue));
  const maxSerial = pivotMaxDateSerial(dateSystem);
  if (!pivotFormulaDateParts(start, dateSystem) || !Number.isFinite(days) || Math.abs(days) > maxSerial) return "#NUM!";
  const holidayResult = pivotHolidaySet(options.holidays || [], dateSystem);
  if (holidayResult.error) return holidayResult.error;
  const weekendResult = pivotWeekendDays(options.weekend ?? 1);
  if (weekendResult.error) return weekendResult.error;
  if (days === 0) return start;
  const target = Math.abs(days);
  const workdaysPerWeek = 7 - weekendResult.weekends.size;
  const estimate = Math.ceil(target / workdaysPerWeek) * 7 + holidayResult.holidays.size + 7;
  if (days > 0) {
    const rangeStart = start + 1;
    let high = Math.min(maxSerial, start + estimate);
    while (high < maxSerial && pivotBusinessDaysBetween(rangeStart, high, holidayResult.holidays, dateSystem, weekendResult.weekends) < target) {
      high = Math.min(maxSerial, high + Math.max(7, high - start));
    }
    if (rangeStart > maxSerial || pivotBusinessDaysBetween(rangeStart, high, holidayResult.holidays, dateSystem, weekendResult.weekends) < target) return "#NUM!";
    let low = rangeStart;
    while (low < high) {
      const middle = Math.floor((low + high) / 2);
      if (pivotBusinessDaysBetween(rangeStart, middle, holidayResult.holidays, dateSystem, weekendResult.weekends) >= target) high = middle;
      else low = middle + 1;
    }
    return low;
  }
  const rangeEnd = start - 1;
  let low = Math.max(0, start - estimate);
  while (low > 0 && pivotBusinessDaysBetween(low, rangeEnd, holidayResult.holidays, dateSystem, weekendResult.weekends) < target) {
    low = Math.max(0, low - Math.max(7, start - low));
  }
  if (rangeEnd < 0 || pivotBusinessDaysBetween(low, rangeEnd, holidayResult.holidays, dateSystem, weekendResult.weekends) < target) return "#NUM!";
  let high = rangeEnd;
  while (low < high) {
    const middle = Math.ceil((low + high) / 2);
    if (pivotBusinessDaysBetween(middle, rangeEnd, holidayResult.holidays, dateSystem, weekendResult.weekends) >= target) low = middle;
    else high = middle - 1;
  }
  return low;
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
