// The workbook evaluator intentionally has a bounded financial profile.  These
// helpers only accept finite numeric cash-flow vectors so an agent never gets a
// plausible-looking answer after a coercion or an unbounded root search.
export const FINANCIAL_MAX_CASH_FLOWS = 10_000;
const FINANCIAL_MAX_RATE_PERIODS = FINANCIAL_MAX_CASH_FLOWS - 1;

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

function fixedCashFlowTerms({ rate, nper, pmt = 0, pv = 0, fv = 0, type = 0 }, helpers) {
  const values = nper === undefined ? [rate, pmt, pv, fv, type] : [rate, nper, pmt, pv, fv, type];
  const safeValues = values.map((value) => finiteNumber(value, helpers));
  const error = safeValues.find((item) => item.error)?.error;
  if (error) return { error };

  const [safeRate, safeNper, safePmt, safePv, safeFv, safeType] = nper === undefined
    ? [safeValues[0], undefined, safeValues[1], safeValues[2], safeValues[3], safeValues[4]]
    : safeValues;
  if (safeRate.value <= -1 || (safeNper && safeNper.value <= 0) || ![0, 1].includes(safeType.value)) return { error: "#NUM!" };
  return {
    rate: safeRate.value,
    ...(safeNper ? { nper: safeNper.value } : {}),
    pmt: safePmt.value,
    pv: safePv.value,
    fv: safeFv.value,
    type: safeType.value,
  };
}

function paymentTerms({ rate, nper, pv, fv = 0, type = 0 }, helpers) {
  return fixedCashFlowTerms({ rate, nper, pv, fv, type }, helpers);
}

function paymentFromTerms({ rate, nper, pv, fv, type }) {
  if (rate === 0) {
    const payment = -(pv + fv) / nper;
    return Number.isFinite(payment) ? payment : "#NUM!";
  }
  const growth = Math.exp(nper * Math.log1p(rate));
  const denominator = (1 + rate * type) * (growth - 1);
  const payment = -(rate * (pv * growth + fv)) / denominator;
  return Number.isFinite(growth) && Number.isFinite(denominator) && denominator !== 0 && Number.isFinite(payment) ? payment : "#NUM!";
}

function paymentPeriod(per, nper, helpers) {
  const safePer = finiteNumber(per, helpers);
  if (safePer.error) return safePer;
  if (!Number.isInteger(safePer.value) || safePer.value < 1 || safePer.value > nper) return { error: "#NUM!" };
  return { value: safePer.value };
}

function cumulativePaymentTerms(rawTerms, helpers) {
  const terms = paymentTerms(rawTerms, helpers);
  if (terms.error) return terms;
  const startPeriod = paymentPeriod(rawTerms.startPeriod, terms.nper, helpers);
  const endPeriod = paymentPeriod(rawTerms.endPeriod, terms.nper, helpers);
  if (startPeriod.error) return startPeriod;
  if (endPeriod.error) return endPeriod;
  if (terms.rate <= 0 || terms.pv <= 0 || endPeriod.value < startPeriod.value || endPeriod.value > FINANCIAL_MAX_RATE_PERIODS) {
    return { error: "#NUM!" };
  }
  return { ...terms, startPeriod: startPeriod.value, endPeriod: endPeriod.value };
}

function depreciationTerms({ cost, salvage, life, period }, helpers) {
  const [safeCost, safeSalvage, safeLife, safePeriod] = [cost, salvage, life, period].map((value) => finiteNumber(value, helpers));
  const error = [safeCost, safeSalvage, safeLife, safePeriod].find((item) => item.error)?.error;
  if (error) return { error };
  if (safeCost.value < 0 || safeSalvage.value < 0 || safeSalvage.value > safeCost.value
    || !Number.isInteger(safeLife.value) || safeLife.value < 1 || safeLife.value > FINANCIAL_MAX_RATE_PERIODS
    || !Number.isInteger(safePeriod.value) || safePeriod.value < 1) return { error: "#NUM!" };
  return { cost: safeCost.value, salvage: safeSalvage.value, life: safeLife.value, period: safePeriod.value };
}

export function calculatePmt(rawTerms, helpers) {
  const terms = paymentTerms(rawTerms, helpers);
  return terms.error ? terms.error : paymentFromTerms(terms);
}

export function calculateIpmt(rawTerms, helpers) {
  const terms = paymentTerms(rawTerms, helpers);
  if (terms.error) return terms.error;
  const safePer = paymentPeriod(rawTerms.per, terms.nper, helpers);
  if (safePer.error) return safePer.error;
  if (terms.rate === 0 || (terms.type === 1 && safePer.value === 1)) return 0;

  const payment = paymentFromTerms(terms);
  if (typeof payment !== "number") return payment;
  const logRate = Math.log1p(terms.rate);
  if (!Number.isFinite(logRate)) return "#NUM!";

  let balance;
  if (terms.type === 0) {
    const elapsedPeriods = safePer.value - 1;
    const growth = Math.exp(elapsedPeriods * logRate);
    const accumulatedPayments = Math.expm1(elapsedPeriods * logRate) / terms.rate;
    balance = terms.pv * growth + payment * accumulatedPayments;
  } else {
    // With payments at the beginning of a period, the first payment has no
    // accrued interest. Later payments service the balance after the prior
    // beginning-of-period payments have accrued for one or more periods.
    const elapsedPeriods = safePer.value - 1;
    const growth = Math.exp(elapsedPeriods * logRate);
    const previousPaymentGrowth = (1 + terms.rate) * Math.expm1((elapsedPeriods - 1) * logRate) / terms.rate;
    balance = (terms.pv + payment) * growth + payment * previousPaymentGrowth;
  }
  const interest = -balance * terms.rate;
  return Number.isFinite(balance) && Number.isFinite(interest) ? interest : "#NUM!";
}

export function calculatePpmt(rawTerms, helpers) {
  const payment = calculatePmt(rawTerms, helpers);
  if (typeof payment !== "number") return payment;
  const interest = calculateIpmt(rawTerms, helpers);
  return typeof interest === "number" ? payment - interest : interest;
}

function calculateCumulativePayment(rawTerms, helpers, calculateComponent) {
  const terms = cumulativePaymentTerms(rawTerms, helpers);
  if (terms.error) return terms.error;
  let total = 0;
  for (let per = terms.startPeriod; per <= terms.endPeriod; per += 1) {
    const component = calculateComponent({ ...terms, per }, helpers);
    if (typeof component !== "number") return component;
    total += component;
    if (!Number.isFinite(total)) return "#NUM!";
  }
  return total;
}

export function calculateCumipmt(rawTerms, helpers) {
  return calculateCumulativePayment(rawTerms, helpers, calculateIpmt);
}

export function calculateCumprinc(rawTerms, helpers) {
  return calculateCumulativePayment(rawTerms, helpers, calculatePpmt);
}

export function calculatePv(rawTerms, helpers) {
  const terms = fixedCashFlowTerms(rawTerms, helpers);
  if (terms.error) return terms.error;
  if (terms.rate === 0) {
    const presentValue = -(terms.pmt * terms.nper + terms.fv);
    return Number.isFinite(presentValue) ? presentValue : "#NUM!";
  }

  const logRate = Math.log1p(terms.rate);
  const growth = Math.exp(terms.nper * logRate);
  const annuity = Math.expm1(terms.nper * logRate) / terms.rate;
  const presentValue = -(terms.fv + terms.pmt * (1 + terms.rate * terms.type) * annuity) / growth;
  return Number.isFinite(logRate) && Number.isFinite(growth) && growth !== 0 && Number.isFinite(annuity) && Number.isFinite(presentValue)
    ? presentValue
    : "#NUM!";
}

export function calculateFv(rawTerms, helpers) {
  const terms = fixedCashFlowTerms(rawTerms, helpers);
  if (terms.error) return terms.error;
  if (terms.rate === 0) {
    const futureValue = -(terms.pv + terms.pmt * terms.nper);
    return Number.isFinite(futureValue) ? futureValue : "#NUM!";
  }

  const logRate = Math.log1p(terms.rate);
  const growth = Math.exp(terms.nper * logRate);
  const annuity = Math.expm1(terms.nper * logRate) / terms.rate;
  const futureValue = -(terms.pv * growth + terms.pmt * (1 + terms.rate * terms.type) * annuity);
  return Number.isFinite(logRate) && Number.isFinite(growth) && Number.isFinite(annuity) && Number.isFinite(futureValue)
    ? futureValue
    : "#NUM!";
}

export function calculateNper(rawTerms, helpers) {
  const terms = fixedCashFlowTerms(rawTerms, helpers);
  if (terms.error) return terms.error;
  if (terms.rate === 0) {
    if (terms.pmt === 0) return "#NUM!";
    const periods = -(terms.pv + terms.fv) / terms.pmt;
    return Number.isFinite(periods) ? periods : "#NUM!";
  }

  const paymentFactor = 1 + terms.rate * terms.type;
  const numerator = terms.pmt * paymentFactor - terms.fv * terms.rate;
  const denominator = terms.pv * terms.rate + terms.pmt * paymentFactor;
  const ratio = numerator / denominator;
  const logRate = Math.log1p(terms.rate);
  const periods = Math.log(ratio) / logRate;
  return Number.isFinite(paymentFactor) && Number.isFinite(numerator) && Number.isFinite(denominator)
    && denominator !== 0 && Number.isFinite(ratio) && ratio > 0 && Number.isFinite(logRate) && logRate !== 0 && Number.isFinite(periods)
    ? periods
    : "#NUM!";
}

export function calculateSln({ cost, salvage, life }, helpers) {
  const [safeCost, safeSalvage, safeLife] = [cost, salvage, life].map((value) => finiteNumber(value, helpers));
  const error = [safeCost, safeSalvage, safeLife].find((item) => item.error)?.error;
  if (error) return error;
  if (safeLife.value === 0) return "#DIV/0!";
  const depreciation = (safeCost.value - safeSalvage.value) / safeLife.value;
  return Number.isFinite(depreciation) ? depreciation : "#NUM!";
}

export function calculateDb({ cost, salvage, life, period, month = 12 }, helpers) {
  const terms = depreciationTerms({ cost, salvage, life, period }, helpers);
  if (terms.error) return terms.error;
  const safeMonth = finiteNumber(month, helpers);
  if (safeMonth.error) return safeMonth.error;
  if (!Number.isInteger(safeMonth.value) || safeMonth.value < 1 || safeMonth.value > 12) return "#NUM!";
  const maxPeriod = terms.life + (safeMonth.value < 12 ? 1 : 0);
  if (terms.period > maxPeriod) return "#NUM!";
  if (terms.cost === 0 || terms.cost === terms.salvage) return 0;

  // Excel/LibreOffice DB uses a three-decimal fixed declining rate. The first
  // and optional final partial years are prorated by month; middle periods use
  // the opening book value from the preceding period without a silent
  // straight-line switch.
  const rate = Math.round((1 - Math.pow(terms.salvage / terms.cost, 1 / terms.life)) * 1_000) / 1_000;
  if (!Number.isFinite(rate) || rate < 0 || rate > 1) return "#NUM!";
  let bookValue = terms.cost;
  for (let currentPeriod = 1; currentPeriod <= terms.period; currentPeriod += 1) {
    const fraction = currentPeriod === 1
      ? safeMonth.value / 12
      : currentPeriod === terms.life + 1 ? (12 - safeMonth.value) / 12 : 1;
    const depreciation = bookValue * rate * fraction;
    if (!Number.isFinite(depreciation)) return "#NUM!";
    if (currentPeriod === terms.period) return depreciation;
    bookValue -= depreciation;
  }
  return "#NUM!";
}

export function calculateDdb({ cost, salvage, life, period, factor = 2 }, helpers) {
  const terms = depreciationTerms({ cost, salvage, life, period }, helpers);
  if (terms.error) return terms.error;
  const safeFactor = finiteNumber(factor, helpers);
  if (safeFactor.error) return safeFactor.error;
  if (safeFactor.value <= 0 || terms.period > terms.life) return "#NUM!";
  if (terms.cost === 0 || terms.cost === terms.salvage) return 0;

  const rate = safeFactor.value / terms.life;
  if (!Number.isFinite(rate)) return "#NUM!";
  let bookValue = terms.cost;
  for (let currentPeriod = 1; currentPeriod <= terms.period; currentPeriod += 1) {
    const depreciation = Math.min(bookValue * rate, Math.max(0, bookValue - terms.salvage));
    if (!Number.isFinite(depreciation)) return "#NUM!";
    if (currentPeriod === terms.period) return depreciation;
    bookValue -= depreciation;
  }
  return "#NUM!";
}

export function calculateRate({ nper, pmt, pv, fv = 0, type = 0, guess }, helpers) {
  const [safeNper, safePmt, safePv, safeFv, safeType] = [nper, pmt, pv, fv, type].map((value) => finiteNumber(value, helpers));
  const error = [safeNper, safePmt, safePv, safeFv, safeType].find((item) => item.error)?.error;
  if (error) return error;
  if (!Number.isInteger(safeNper.value) || safeNper.value < 1 || safeNper.value > FINANCIAL_MAX_RATE_PERIODS || ![0, 1].includes(safeType.value)) return "#NUM!";
  const safeGuess = validGuess(guess, helpers);
  if (safeGuess.error) return safeGuess.error;

  // RATE solves the same present/future-value equation as PMT.  Expressing
  // it as dated cash flows lets it reuse the bounded log-rate root search and
  // preserves the distinct beginning-of-period payment timing semantics.
  const cashFlows = Array(safeNper.value + 1).fill(safePmt.value);
  if (safeType.value === 0) {
    cashFlows[0] = safePv.value;
    cashFlows[safeNper.value] += safeFv.value;
  } else {
    cashFlows[0] = safePv.value + safePmt.value;
    cashFlows[safeNper.value] = safeFv.value;
  }
  return returnRate(cashFlows, cashFlows.map((_, index) => index), safeGuess.value);
}

export function calculateNpv({ rate, cashFlows }, helpers) {
  const safeRate = validRate(rate, helpers);
  if (safeRate.error) return safeRate.error;
  const safeCashFlows = finiteSeries(cashFlows, helpers);
  if (safeCashFlows.error) return safeCashFlows.error;
  return discountedSum(safeCashFlows.values, safeCashFlows.values.map((_, index) => index + 1), safeRate.value);
}

export function calculateMirr({ cashFlows, financeRate, reinvestRate }, helpers) {
  const safeCashFlows = finiteSeries(cashFlows, helpers);
  if (safeCashFlows.error) return safeCashFlows.error;
  if (safeCashFlows.values.length < 2 || !hasReturnSignPattern(safeCashFlows.values)) return "#NUM!";
  const safeFinanceRate = validRate(financeRate, helpers);
  if (safeFinanceRate.error) return safeFinanceRate.error;
  const safeReinvestRate = validRate(reinvestRate, helpers);
  if (safeReinvestRate.error) return safeReinvestRate.error;

  const finalPeriod = safeCashFlows.values.length - 1;
  const financeLogRate = Math.log1p(safeFinanceRate.value);
  const reinvestLogRate = Math.log1p(safeReinvestRate.value);
  if (!Number.isFinite(financeLogRate) || !Number.isFinite(reinvestLogRate)) return "#NUM!";

  let negativePresentValue = 0;
  let positiveFutureValue = 0;
  for (let index = 0; index <= finalPeriod; index += 1) {
    const cashFlow = safeCashFlows.values[index];
    if (cashFlow < 0) {
      const discount = Math.exp(-index * financeLogRate);
      const presentValue = cashFlow * discount;
      if (!Number.isFinite(discount) || !Number.isFinite(presentValue)) return "#NUM!";
      negativePresentValue += presentValue;
    } else if (cashFlow > 0) {
      const growth = Math.exp((finalPeriod - index) * reinvestLogRate);
      const futureValue = cashFlow * growth;
      if (!Number.isFinite(growth) || !Number.isFinite(futureValue)) return "#NUM!";
      positiveFutureValue += futureValue;
    }
    if (!Number.isFinite(negativePresentValue) || !Number.isFinite(positiveFutureValue)) return "#NUM!";
  }

  const ratio = -positiveFutureValue / negativePresentValue;
  const result = Math.pow(ratio, 1 / finalPeriod) - 1;
  return Number.isFinite(ratio) && ratio > 0 && Number.isFinite(result) ? result : "#NUM!";
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
