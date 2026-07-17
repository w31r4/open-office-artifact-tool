// The workbook evaluator intentionally has a bounded financial profile.  These
// helpers only accept finite numeric cash-flow vectors so an agent never gets a
// plausible-looking answer after a coercion or an unbounded root search.
export const FINANCIAL_MAX_CASH_FLOWS = 10_000;

const ROOT_LOG_RATE_MIN = -20;
const ROOT_LOG_RATE_MAX = 20;
const ROOT_GRID_STEPS = 384;
const ROOT_MAX_ITERATIONS = 128;
const ROOT_VALUE_TOLERANCE = 1e-10;
const ROOT_LOG_RATE_TOLERANCE = 1e-12;

function finiteNumber(value, helpers) {
  const error = helpers.errorCode(value);
  if (error) return { error };
  if (typeof value !== "number" || !Number.isFinite(value)) return { error: "#VALUE!" };
  return { value };
}

function finiteSeries(values, helpers) {
  if (!Array.isArray(values) || values.length === 0) return { error: "#VALUE!" };
  if (values.length > FINANCIAL_MAX_CASH_FLOWS) return { error: "#NUM!" };
  const numbers = [];
  for (const value of values) {
    const number = finiteNumber(value, helpers);
    if (number.error) return number;
    numbers.push(number.value);
  }
  return { values: numbers };
}

function dateSeries(values, helpers) {
  if (!Array.isArray(values) || values.length === 0) return { error: "#VALUE!" };
  if (values.length > FINANCIAL_MAX_CASH_FLOWS) return { error: "#NUM!" };
  const serials = [];
  for (const value of values) {
    const serialValue = helpers.dateNumber(value);
    const error = helpers.errorCode(serialValue);
    if (error) return { error };
    if (!Number.isFinite(serialValue)) return { error: "#VALUE!" };
    const serial = Math.floor(serialValue);
    if (!helpers.isValidDate(serial)) return { error: "#NUM!" };
    serials.push(serial);
  }
  return { values: serials };
}

function validRate(rawRate, helpers) {
  const rate = finiteNumber(rawRate, helpers);
  if (rate.error) return rate;
  return rate.value > -1 ? rate : { error: "#NUM!" };
}

function validGuess(rawGuess, helpers) {
  if (rawGuess === undefined) return { value: 0.1 };
  return validRate(rawGuess, helpers);
}

function discountedSum(cashFlows, exponents, rate) {
  const logRate = Math.log1p(rate);
  if (!Number.isFinite(logRate)) return "#NUM!";
  let total = 0;
  for (let index = 0; index < cashFlows.length; index += 1) {
    const discount = Math.exp(-exponents[index] * logRate);
    const term = cashFlows[index] * discount;
    if (!Number.isFinite(discount) || !Number.isFinite(term)) return "#NUM!";
    total += term;
    if (!Number.isFinite(total)) return "#NUM!";
  }
  return total;
}

function hasReturnSignPattern(cashFlows) {
  return cashFlows.some((value) => value < 0) && cashFlows.some((value) => value > 0);
}

function normalizedCashFlows(cashFlows) {
  let maximum = 0;
  for (const value of cashFlows) maximum = Math.max(maximum, Math.abs(value));
  return maximum > 0 ? cashFlows.map((value) => value / maximum) : undefined;
}

function evaluateAtLogRate(cashFlows, exponents, logRate) {
  let value = 0;
  let derivative = 0;
  for (let index = 0; index < cashFlows.length; index += 1) {
    const exponent = exponents[index];
    const power = Math.exp(-exponent * logRate);
    const term = cashFlows[index] * power;
    if (!Number.isFinite(power) || !Number.isFinite(term)) {
      if (!Number.isFinite(term)) return { value: Math.sign(term) * Infinity, derivative: Math.sign(term) * Infinity };
      return undefined;
    }
    value += term;
    derivative -= exponent * term;
    if (!Number.isFinite(value) || !Number.isFinite(derivative)) return { value: Math.sign(value) * Infinity, derivative: Math.sign(derivative) * Infinity };
  }
  return { value, derivative };
}

function sign(value) {
  return value > 0 ? 1 : value < 0 ? -1 : 0;
}

function rateFromLogRate(logRate) {
  const rate = Math.expm1(logRate);
  return Number.isFinite(rate) && rate > -1 ? rate : undefined;
}

function newtonRoot(evaluate, initialLogRate) {
  let logRate = initialLogRate;
  for (let iteration = 0; iteration < ROOT_MAX_ITERATIONS; iteration += 1) {
    const result = evaluate(logRate);
    if (!result || !Number.isFinite(result.value) || !Number.isFinite(result.derivative)) return undefined;
    if (Math.abs(result.value) <= ROOT_VALUE_TOLERANCE) return rateFromLogRate(logRate);
    if (Math.abs(result.derivative) <= Number.EPSILON) return undefined;
    const next = logRate - result.value / result.derivative;
    if (!Number.isFinite(next) || next <= ROOT_LOG_RATE_MIN || next >= ROOT_LOG_RATE_MAX) return undefined;
    if (Math.abs(next - logRate) <= ROOT_LOG_RATE_TOLERANCE) return rateFromLogRate(next);
    logRate = next;
  }
  return undefined;
}

function bracketedRoot(evaluate, initialLogRate) {
  const grid = Array.from(
    { length: ROOT_GRID_STEPS + 1 },
    (_, index) => ROOT_LOG_RATE_MIN + (ROOT_LOG_RATE_MAX - ROOT_LOG_RATE_MIN) * index / ROOT_GRID_STEPS,
  );
  if (initialLogRate > ROOT_LOG_RATE_MIN && initialLogRate < ROOT_LOG_RATE_MAX) grid.push(initialLogRate);
  grid.sort((left, right) => left - right);

  let bestExact;
  const brackets = [];
  let previousLogRate;
  let previous;
  for (const logRate of grid) {
    const current = evaluate(logRate);
    if (!current || Number.isNaN(current.value)) {
      previousLogRate = undefined;
      previous = undefined;
      continue;
    }
    if (current.value === 0 || (Number.isFinite(current.value) && Math.abs(current.value) <= ROOT_VALUE_TOLERANCE)) {
      if (bestExact == null || Math.abs(logRate - initialLogRate) < Math.abs(bestExact - initialLogRate)) bestExact = logRate;
    }
    if (previous && sign(previous.value) && sign(current.value) && sign(previous.value) !== sign(current.value)) {
      brackets.push({ left: previousLogRate, right: logRate });
    }
    previousLogRate = logRate;
    previous = current;
  }
  if (bestExact != null) return rateFromLogRate(bestExact);
  if (!brackets.length) return undefined;

  brackets.sort((left, right) => Math.abs((left.left + left.right) / 2 - initialLogRate) - Math.abs((right.left + right.right) / 2 - initialLogRate));
  let { left, right } = brackets[0];
  let leftValue = evaluate(left)?.value;
  const rightValue = evaluate(right)?.value;
  if (!sign(leftValue) || !sign(rightValue) || sign(leftValue) === sign(rightValue)) return undefined;

  for (let iteration = 0; iteration < ROOT_MAX_ITERATIONS; iteration += 1) {
    const middle = (left + right) / 2;
    const middleValue = evaluate(middle)?.value;
    if (middleValue == null || Number.isNaN(middleValue)) return undefined;
    if (middleValue === 0 || (Number.isFinite(middleValue) && Math.abs(middleValue) <= ROOT_VALUE_TOLERANCE)) return rateFromLogRate(middle);
    if (Math.abs(right - left) <= ROOT_LOG_RATE_TOLERANCE) return rateFromLogRate(middle);
    if (sign(leftValue) !== sign(middleValue)) {
      right = middle;
    } else {
      left = middle;
      leftValue = middleValue;
    }
  }
  return rateFromLogRate((left + right) / 2);
}

function returnRate(cashFlows, exponents, guess) {
  if (!hasReturnSignPattern(cashFlows)) return "#NUM!";
  const normalized = normalizedCashFlows(cashFlows);
  if (!normalized) return "#NUM!";
  const initialLogRate = Math.log1p(guess);
  if (!Number.isFinite(initialLogRate) || initialLogRate < ROOT_LOG_RATE_MIN || initialLogRate > ROOT_LOG_RATE_MAX) return "#NUM!";
  const evaluate = (logRate) => evaluateAtLogRate(normalized, exponents, logRate);
  return newtonRoot(evaluate, initialLogRate) ?? bracketedRoot(evaluate, initialLogRate) ?? "#NUM!";
}

export function calculatePmt({ rate, nper, pv, fv = 0, type = 0 }, helpers) {
  const [safeRate, safeNper, safePv, safeFv, safeType] = [rate, nper, pv, fv, type].map((value) => finiteNumber(value, helpers));
  const error = [safeRate, safeNper, safePv, safeFv, safeType].find((item) => item.error)?.error;
  if (error) return error;
  if (safeRate.value <= -1 || safeNper.value <= 0 || ![0, 1].includes(safeType.value)) return "#NUM!";
  if (safeRate.value === 0) {
    const payment = -(safePv.value + safeFv.value) / safeNper.value;
    return Number.isFinite(payment) ? payment : "#NUM!";
  }
  const growth = Math.exp(safeNper.value * Math.log1p(safeRate.value));
  const denominator = (1 + safeRate.value * safeType.value) * (growth - 1);
  const payment = -(safeRate.value * (safePv.value * growth + safeFv.value)) / denominator;
  return Number.isFinite(growth) && Number.isFinite(denominator) && denominator !== 0 && Number.isFinite(payment) ? payment : "#NUM!";
}

export function calculateNpv({ rate, cashFlows }, helpers) {
  const safeRate = validRate(rate, helpers);
  if (safeRate.error) return safeRate.error;
  const safeCashFlows = finiteSeries(cashFlows, helpers);
  if (safeCashFlows.error) return safeCashFlows.error;
  return discountedSum(safeCashFlows.values, safeCashFlows.values.map((_, index) => index + 1), safeRate.value);
}

export function calculateXnpv({ rate, cashFlows, dates }, helpers) {
  const safeRate = validRate(rate, helpers);
  if (safeRate.error) return safeRate.error;
  const safeCashFlows = finiteSeries(cashFlows, helpers);
  if (safeCashFlows.error) return safeCashFlows.error;
  const safeDates = dateSeries(dates, helpers);
  if (safeDates.error) return safeDates.error;
  if (safeCashFlows.values.length !== safeDates.values.length) return "#VALUE!";
  const baseDate = safeDates.values[0];
  return discountedSum(safeCashFlows.values, safeDates.values.map((date) => (date - baseDate) / 365), safeRate.value);
}

export function calculateIrr({ cashFlows, guess }, helpers) {
  const safeCashFlows = finiteSeries(cashFlows, helpers);
  if (safeCashFlows.error) return safeCashFlows.error;
  const safeGuess = validGuess(guess, helpers);
  if (safeGuess.error) return safeGuess.error;
  return returnRate(safeCashFlows.values, safeCashFlows.values.map((_, index) => index), safeGuess.value);
}

export function calculateXirr({ cashFlows, dates, guess }, helpers) {
  const safeCashFlows = finiteSeries(cashFlows, helpers);
  if (safeCashFlows.error) return safeCashFlows.error;
  const safeDates = dateSeries(dates, helpers);
  if (safeDates.error) return safeDates.error;
  if (safeCashFlows.values.length !== safeDates.values.length) return "#VALUE!";
  const safeGuess = validGuess(guess, helpers);
  if (safeGuess.error) return safeGuess.error;
  const baseDate = safeDates.values[0];
  return returnRate(safeCashFlows.values, safeDates.values.map((date) => (date - baseDate) / 365), safeGuess.value);
}
