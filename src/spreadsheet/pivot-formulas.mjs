import { formulaTimeParts, formulaTimeSerial } from "./formula-coercion.mjs";
import { pivotDateParts, pivotDateSerial, pivotFormulaDateParts, pivotShiftDateMonths, pivotWeekdayIndex } from "./pivot-dates.mjs";

const PIVOT_FUNCTIONS = new Map([
  ["ABS", { minArgs: 1, maxArgs: 1 }],
  ["SUM", { minArgs: 1, maxArgs: 32 }],
  ["MIN", { minArgs: 1, maxArgs: 32 }],
  ["MAX", { minArgs: 1, maxArgs: 32 }],
  ["AVERAGE", { minArgs: 1, maxArgs: 32 }],
  ["PRODUCT", { minArgs: 1, maxArgs: 32 }],
  ["ROUND", { minArgs: 2, maxArgs: 2 }],
  ["POWER", { minArgs: 2, maxArgs: 2 }],
  ["SQRT", { minArgs: 1, maxArgs: 1 }],
  ["MOD", { minArgs: 2, maxArgs: 2 }],
  ["SIGN", { minArgs: 1, maxArgs: 1 }],
  ["INT", { minArgs: 1, maxArgs: 1 }],
  ["AND", { minArgs: 1, maxArgs: 32 }],
  ["OR", { minArgs: 1, maxArgs: 32 }],
  ["NOT", { minArgs: 1, maxArgs: 1 }],
  ["IF", { minArgs: 2, maxArgs: 3 }],
  ["IFERROR", { minArgs: 2, maxArgs: 2 }],
  ["IFNA", { minArgs: 2, maxArgs: 2 }],
  ["NA", { minArgs: 0, maxArgs: 0 }],
  ["ISERROR", { minArgs: 1, maxArgs: 1 }],
  ["ISNUMBER", { minArgs: 1, maxArgs: 1 }],
  ["ISTEXT", { minArgs: 1, maxArgs: 1 }],
  ["LEN", { minArgs: 1, maxArgs: 1 }],
  ["LEFT", { minArgs: 1, maxArgs: 2 }],
  ["RIGHT", { minArgs: 1, maxArgs: 2 }],
  ["MID", { minArgs: 3, maxArgs: 3 }],
  ["LOWER", { minArgs: 1, maxArgs: 1 }],
  ["UPPER", { minArgs: 1, maxArgs: 1 }],
  ["TRIM", { minArgs: 1, maxArgs: 1 }],
  ["DATE", { minArgs: 3, maxArgs: 3 }],
  ["YEAR", { minArgs: 1, maxArgs: 1 }],
  ["MONTH", { minArgs: 1, maxArgs: 1 }],
  ["DAY", { minArgs: 1, maxArgs: 1 }],
  ["EDATE", { minArgs: 2, maxArgs: 2 }],
  ["EOMONTH", { minArgs: 2, maxArgs: 2 }],
  ["DAYS", { minArgs: 2, maxArgs: 2 }],
  ["WEEKDAY", { minArgs: 1, maxArgs: 2 }],
  ["TIME", { minArgs: 3, maxArgs: 3 }],
  ["HOUR", { minArgs: 1, maxArgs: 1 }],
  ["MINUTE", { minArgs: 1, maxArgs: 1 }],
  ["SECOND", { minArgs: 1, maxArgs: 1 }],
]);

const COMPARISON_OPERATORS = new Set(["=", "<>", "<", "<=", ">", ">="]);

function quotedStringToken(text, index) {
  let value = "";
  index += 1;
  while (index < text.length) {
    if (text[index] !== '"') { value += text[index++]; continue; }
    if (text[index + 1] === '"') { value += '"'; index += 2; continue; }
    return { token: { type: "string", value }, nextIndex: index + 1 };
  }
  throw new SyntaxError("PivotTable calculated field formula has an unterminated string constant.");
}

function formulaTokens(formula, fields) {
  const text = String(formula || "").trim().replace(/^=/, "");
  if (!text) throw new TypeError("PivotTable calculated field formula must not be empty.");
  if (text.length > 4096) throw new RangeError("PivotTable calculated field formula exceeds 4096 characters.");
  const known = new Set(fields);
  const tokens = [];
  for (let index = 0; index < text.length;) {
    if (/\s/.test(text[index])) { index += 1; continue; }
    const rest = text.slice(index);
    const number = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?/.exec(rest)?.[0];
    if (number) { tokens.push({ type: "number", value: Number(number) }); index += number.length; continue; }
    if (text[index] === '"') {
      const parsed = quotedStringToken(text, index);
      tokens.push(parsed.token);
      index = parsed.nextIndex;
      continue;
    }
    const comparison = /^(?:<>|<=|>=)/.exec(rest)?.[0];
    if (comparison) { tokens.push({ type: "operator", value: comparison }); index += comparison.length; continue; }
    if ("+-*/^()%,&=<>".includes(text[index])) { tokens.push({ type: "operator", value: text[index++] }); continue; }
    let field;
    if (text[index] === "[") {
      const end = text.indexOf("]", index + 1);
      if (end < 0) throw new SyntaxError("PivotTable calculated field formula has an unterminated [field] reference.");
      field = text.slice(index + 1, end).trim();
      index = end + 1;
    } else if (text[index] === "'") {
      let value = "";
      index += 1;
      while (index < text.length) {
        if (text[index] !== "'") { value += text[index++]; continue; }
        if (text[index + 1] === "'") { value += "'"; index += 2; continue; }
        index += 1;
        field = value;
        break;
      }
      if (field == null) throw new SyntaxError("PivotTable calculated field formula has an unterminated quoted field reference.");
    } else {
      const identifier = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(rest)?.[0];
      if (!identifier) throw new SyntaxError(`PivotTable calculated field formula has unsupported token near ${rest.slice(0, 12)}.`);
      const next = text.slice(index + identifier.length).trimStart()[0];
      const normalized = identifier.toUpperCase();
      if (next === "(") {
        if (!PIVOT_FUNCTIONS.has(normalized)) throw new SyntaxError(`PivotTable calculated field formula uses unsupported function ${identifier}.`);
        tokens.push({ type: "function", value: normalized });
        index += identifier.length;
        continue;
      }
      if (normalized === "TRUE" || normalized === "FALSE") {
        tokens.push({ type: "boolean", value: normalized === "TRUE" });
        index += identifier.length;
        continue;
      }
      field = identifier;
      index += identifier.length;
    }
    if (!known.has(field)) throw new Error(`PivotTable calculated field formula references unknown source field ${field}.`);
    tokens.push({ type: "field", value: field });
  }
  if (tokens.length > 512) throw new RangeError("PivotTable calculated field formula exceeds 512 tokens.");
  return tokens;
}

function renderPivotFormula(tokens) {
  return tokens.map((token) => {
    if (token.type === "field") return `'${token.value.replaceAll("'", "''")}'`;
    if (token.type === "string") return `"${token.value.replaceAll('"', '""')}"`;
    if (token.type === "boolean") return token.value ? "TRUE" : "FALSE";
    return String(token.value);
  }).join("");
}

function uniqueReferences(tokens) {
  return [...new Set(tokens.filter((token) => token.type === "field").map((token) => token.value))];
}

function pivotFormulaAst(tokens) {
  let index = 0;
  const primary = () => {
    const token = tokens[index++];
    if (!token) throw new SyntaxError("PivotTable calculated field formula ended unexpectedly.");
    if (["number", "string", "boolean"].includes(token.type)) return { kind: "literal", value: token.value };
    if (token.type === "field") return { kind: "field", name: token.value };
    if (token.type === "function") {
      if (tokens[index++]?.value !== "(") throw new SyntaxError(`PivotTable ${token.value} requires an opening parenthesis.`);
      const args = [];
      if (tokens[index]?.value !== ")") {
        args.push(comparison());
        while (tokens[index]?.value === ",") { index += 1; args.push(comparison()); }
      }
      if (tokens[index++]?.value !== ")") throw new SyntaxError(`PivotTable ${token.value} requires a closing parenthesis.`);
      return { kind: "call", name: token.value, args };
    }
    if (token.value === "(") {
      const value = comparison();
      if (tokens[index++]?.value !== ")") throw new SyntaxError("PivotTable calculated field formula requires a closing parenthesis.");
      return value;
    }
    throw new SyntaxError(`PivotTable calculated field formula has unexpected token ${token.value}.`);
  };
  const unary = () => ["+", "-"].includes(tokens[index]?.value) ? { kind: "unary", operator: tokens[index++].value, value: unary() } : primary();
  const power = () => {
    let left = unary();
    if (tokens[index]?.value === "^") { index += 1; left = { kind: "binary", operator: "^", left, right: power() }; }
    return left;
  };
  const percent = () => {
    let value = power();
    while (tokens[index]?.value === "%") { index += 1; value = { kind: "percent", value }; }
    return value;
  };
  const term = () => {
    let left = percent();
    while (["*", "/"].includes(tokens[index]?.value)) { const operator = tokens[index++].value; left = { kind: "binary", operator, left, right: percent() }; }
    return left;
  };
  const additive = () => {
    let left = term();
    while (["+", "-"].includes(tokens[index]?.value)) { const operator = tokens[index++].value; left = { kind: "binary", operator, left, right: term() }; }
    return left;
  };
  const concatenate = () => {
    let left = additive();
    while (tokens[index]?.value === "&") { index += 1; left = { kind: "binary", operator: "&", left, right: additive() }; }
    return left;
  };
  const comparison = () => {
    let left = concatenate();
    while (COMPARISON_OPERATORS.has(tokens[index]?.value)) { const operator = tokens[index++].value; left = { kind: "binary", operator, left, right: concatenate() }; }
    return left;
  };
  const root = comparison();
  if (index !== tokens.length) throw new SyntaxError(`PivotTable calculated field formula has unexpected token ${tokens[index].value}.`);
  return root;
}

function validateFunction(node) {
  const contract = PIVOT_FUNCTIONS.get(node.name);
  if (node.args.length >= contract.minArgs && node.args.length <= contract.maxArgs) return;
  if (contract.minArgs === contract.maxArgs) {
    const count = contract.minArgs === 1 ? "one" : contract.minArgs === 2 ? "two" : String(contract.minArgs);
    throw new SyntaxError(`PivotTable ${node.name} requires exactly ${count} argument${contract.minArgs === 1 ? "" : "s"}.`);
  }
  if (contract.minArgs === 1 && node.args.length < 1) throw new SyntaxError(`PivotTable ${node.name} requires at least one argument.`);
  if (contract.maxArgs === 32 && node.args.length > 32) throw new RangeError(`PivotTable ${node.name} exceeds 32 arguments.`);
  throw new SyntaxError(`PivotTable ${node.name} requires ${contract.minArgs} or ${contract.maxArgs} arguments.`);
}

function validatePivotFormulaAst(node) {
  if (node.kind === "call") validateFunction(node);
  if (node.left) validatePivotFormulaAst(node.left);
  if (node.right) validatePivotFormulaAst(node.right);
  if (node.value?.kind) validatePivotFormulaAst(node.value);
  for (const arg of node.args || []) validatePivotFormulaAst(arg);
}

function parsedPivotFormula(formula, fields) {
  const tokens = formulaTokens(formula, fields);
  const ast = pivotFormulaAst(tokens);
  validatePivotFormulaAst(ast);
  return { tokens, ast };
}

export function pivotFormulaToOoxml(formula, fields) {
  return renderPivotFormula(parsedPivotFormula(formula, fields).tokens);
}

export function normalizeCalculatedFields(value, sourceFields, allowUnsupported = false, options = {}) {
  if (value == null) return [];
  if (!Array.isArray(value)) throw new TypeError("PivotTable calculatedFields must be an array.");
  if (value.length > 128) throw new RangeError("PivotTable calculatedFields exceeds 128 fields.");
  const names = new Set(sourceFields);
  return value.map((entry, index) => {
    const name = String(entry?.name || entry?.field || "").trim();
    if (!name) throw new TypeError(`PivotTable calculatedFields[${index}] requires name.`);
    if (names.has(name)) throw new Error(`PivotTable calculated field name ${name} must be unique and must not replace a source field.`);
    names.add(name);
    const formula = String(entry?.formula || "").trim();
    const numFmtId = entry?.numFmtId == null ? 0 : Number(entry.numFmtId);
    if (!Number.isInteger(numFmtId) || numFmtId < 0) throw new TypeError(`PivotTable calculatedFields[${index}] numFmtId must be a non-negative integer.`);
    try {
      const parsed = parsedPivotFormula(formula, sourceFields);
      evaluatePivotFormula(formula, Object.fromEntries(sourceFields.map((field) => [field, 0])), sourceFields, options);
      return { name, formula: `=${renderPivotFormula(parsed.tokens)}`, numFmtId, references: uniqueReferences(parsed.tokens) };
    } catch (error) {
      if (!allowUnsupported || error instanceof RangeError) throw error;
      return { name, formula: formula.startsWith("=") ? formula : `=${formula}`, numFmtId, references: [], supported: false, error: error.message };
    }
  });
}

function formulaError(value) {
  return typeof value === "string" && /^(?:#NULL!|#DIV\/0!|#VALUE!|#REF!|#NAME\?|#NUM!|#N\/A|#GETTING_DATA|#SPILL!|#CALC!)$/.test(value);
}

const PIVOT_TEXT = Symbol("pivotFormulaText");

function formulaText(value) {
  return { [PIVOT_TEXT]: true, value: String(value) };
}

function textFormulaValue(value) {
  return Boolean(value?.[PIVOT_TEXT]);
}

function scalarFormulaValue(value) {
  return textFormulaValue(value) ? value.value : value;
}

function numericValue(value) {
  if (formulaError(value)) return value;
  const number = Number(scalarFormulaValue(value));
  return Number.isFinite(number) ? number : "#VALUE!";
}

function comparisonValue(left, operator, right) {
  left = scalarFormulaValue(left);
  right = scalarFormulaValue(right);
  const leftNumber = typeof left !== "string" || left.trim() !== "" ? Number(left) : Number.NaN;
  const rightNumber = typeof right !== "string" || right.trim() !== "" ? Number(right) : Number.NaN;
  const numeric = Number.isFinite(leftNumber) && Number.isFinite(rightNumber);
  const a = numeric ? leftNumber : String(left).toLocaleLowerCase("en-US");
  const b = numeric ? rightNumber : String(right).toLocaleLowerCase("en-US");
  if (operator === "=") return a === b;
  if (operator === "<>") return a !== b;
  if (operator === "<") return a < b;
  if (operator === "<=") return a <= b;
  if (operator === ">") return a > b;
  return a >= b;
}

function formulaBinary(left, operator, right) {
  if (formulaError(left)) return left;
  if (formulaError(right)) return right;
  if (COMPARISON_OPERATORS.has(operator)) return comparisonValue(left, operator, right);
  if (operator === "&") return formulaText(`${scalarFormulaValue(left)}${scalarFormulaValue(right)}`);
  const a = numericValue(left);
  const b = numericValue(right);
  if (formulaError(a)) return a;
  if (formulaError(b)) return b;
  if (operator === "/" && b === 0) return "#DIV/0!";
  const result = operator === "+" ? a + b
    : operator === "-" ? a - b
      : operator === "*" ? a * b
        : operator === "/" ? a / b
          : a ** b;
  return Number.isFinite(result) ? result : "#NUM!";
}

function formulaUnary(value, transform) {
  if (formulaError(value)) return value;
  const number = numericValue(value);
  return formulaError(number) ? number : transform(number);
}

function pivotTextValue(value) {
  if (formulaError(value)) return value;
  value = scalarFormulaValue(value);
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

function pivotTextCharacters(value) {
  const text = pivotTextValue(value);
  return formulaError(text) ? text : Array.from(text);
}

function pivotTextCount(value, fallback) {
  if (value == null) return fallback;
  const count = numericValue(value);
  if (formulaError(count)) return count;
  return Math.trunc(count);
}

function pivotTimeParts(value, dateSystem) {
  const parts = formulaTimeParts(scalarFormulaValue(value));
  if (!parts || (parts.dateText && !pivotDateParts(parts.dateText, dateSystem))) return undefined;
  return parts;
}

function formulaFunction(name, args, options) {
  if (name === "ISERROR") return formulaError(args[0]);
  if (name === "ISNUMBER") return typeof args[0] === "number" && Number.isFinite(args[0]);
  if (name === "ISTEXT") return textFormulaValue(args[0]);
  if (name === "NA") return "#N/A";
  const error = args.find(formulaError);
  if (error) return error;
  if (name === "LEN") {
    const characters = pivotTextCharacters(args[0]);
    return formulaError(characters) ? characters : characters.length;
  }
  if (name === "LEFT" || name === "RIGHT") {
    const characters = pivotTextCharacters(args[0]);
    if (formulaError(characters)) return characters;
    const count = pivotTextCount(args[1], 1);
    if (formulaError(count)) return count;
    if (count < 0) return "#VALUE!";
    const result = name === "LEFT" ? characters.slice(0, count) : count === 0 ? [] : characters.slice(-count);
    return formulaText(result.join(""));
  }
  if (name === "MID") {
    const characters = pivotTextCharacters(args[0]);
    if (formulaError(characters)) return characters;
    const start = pivotTextCount(args[1]);
    const count = pivotTextCount(args[2]);
    if (formulaError(start)) return start;
    if (formulaError(count)) return count;
    if (start < 1 || count < 0) return "#VALUE!";
    return formulaText(characters.slice(start - 1, start - 1 + count).join(""));
  }
  if (name === "LOWER" || name === "UPPER" || name === "TRIM") {
    const text = pivotTextValue(args[0]);
    if (formulaError(text)) return text;
    if (name === "TRIM") return formulaText(text.replace(/ +/g, " ").replace(/^ | $/g, ""));
    return formulaText(name === "LOWER" ? text.toLowerCase() : text.toUpperCase());
  }
  if (name === "DATE") {
    const values = args.map(numericValue);
    const valueError = values.find(formulaError);
    if (valueError) return valueError;
    return pivotDateSerial(values[0], values[1], values[2], options.dateSystem) ?? "#NUM!";
  }
  if (name === "YEAR" || name === "MONTH" || name === "DAY") {
    const serial = numericValue(args[0]);
    if (formulaError(serial)) return serial;
    const parts = pivotFormulaDateParts(serial, options.dateSystem);
    return parts ? parts[name.toLowerCase()] : "#NUM!";
  }
  if (name === "EDATE" || name === "EOMONTH") {
    const values = args.map(numericValue);
    const valueError = values.find(formulaError);
    if (valueError) return valueError;
    return pivotShiftDateMonths(values[0], values[1], name === "EOMONTH", options.dateSystem) ?? "#NUM!";
  }
  if (name === "DAYS") {
    const values = args.map(numericValue);
    const valueError = values.find(formulaError);
    if (valueError) return valueError;
    if (values.some((value) => !pivotFormulaDateParts(value, options.dateSystem))) return "#NUM!";
    return Math.floor(values[0]) - Math.floor(values[1]);
  }
  if (name === "WEEKDAY") {
    const serial = numericValue(args[0]);
    const returnTypeValue = args[1] == null ? 1 : numericValue(args[1]);
    if (formulaError(serial)) return serial;
    if (formulaError(returnTypeValue)) return returnTypeValue;
    const weekday = pivotWeekdayIndex(serial, options.dateSystem);
    const returnType = Math.trunc(returnTypeValue);
    if (weekday == null) return "#NUM!";
    if (returnType === 1) return weekday + 1;
    if (returnType === 2 || returnType === 11) return (weekday + 6) % 7 + 1;
    if (returnType === 3) return (weekday + 6) % 7;
    if (returnType >= 12 && returnType <= 17) return (weekday - (returnType - 10) + 7) % 7 + 1;
    return "#NUM!";
  }
  if (name === "TIME") {
    const values = args.map(numericValue);
    const valueError = values.find(formulaError);
    if (valueError) return valueError;
    return formulaTimeSerial(values[0], values[1], values[2]) ?? "#NUM!";
  }
  if (name === "HOUR" || name === "MINUTE" || name === "SECOND") {
    const parts = pivotTimeParts(args[0], options.dateSystem);
    return parts ? parts[name.toLowerCase()] : "#VALUE!";
  }
  if (name === "ABS") return formulaUnary(args[0], Math.abs);
  if (name === "SQRT") {
    const value = numericValue(args[0]);
    if (formulaError(value)) return value;
    return value < 0 ? "#NUM!" : Math.sqrt(value);
  }
  if (name === "SIGN") return formulaUnary(args[0], (value) => value === 0 ? 0 : Math.sign(value));
  if (name === "INT") return formulaUnary(args[0], (value) => Math.floor(value) || 0);
  if (name === "NOT") {
    const condition = formulaCondition(args[0]);
    return formulaError(condition) ? condition : !condition;
  }
  if (name === "AND" || name === "OR") {
    const conditions = args.map(formulaCondition);
    const conditionError = conditions.find(formulaError);
    if (conditionError) return conditionError;
    return name === "AND" ? conditions.every(Boolean) : conditions.some(Boolean);
  }
  if (name === "ROUND") {
    const digits = numericValue(args[1]);
    if (formulaError(digits)) return digits;
    if (!Number.isInteger(digits) || digits < -15 || digits > 15) throw new RangeError("PivotTable ROUND digits must be an integer from -15 to 15.");
    const value = numericValue(args[0]);
    if (formulaError(value)) return value;
    const factor = 10 ** Math.abs(digits);
    const scaled = digits >= 0 ? Math.abs(value) * factor : Math.abs(value) / factor;
    const rounded = Math.sign(value) * Math.round(scaled + Number.EPSILON * scaled) * (digits >= 0 ? 1 / factor : factor);
    return rounded === 0 ? 0 : rounded;
  }
  const numbers = args.map(numericValue);
  const numericError = numbers.find(formulaError);
  if (numericError) return numericError;
  if (name === "SUM") {
    const result = numbers.reduce((sum, value) => sum + value, 0);
    return Number.isFinite(result) ? result : "#NUM!";
  }
  if (name === "MIN") return Math.min(...numbers);
  if (name === "MAX") return Math.max(...numbers);
  if (name === "PRODUCT") {
    const result = numbers.reduce((product, value) => product * value, 1);
    return Number.isFinite(result) ? result : "#NUM!";
  }
  if (name === "POWER") {
    if (numbers[0] === 0 && numbers[1] < 0) return "#DIV/0!";
    const result = numbers[0] ** numbers[1];
    return Number.isFinite(result) ? result : "#NUM!";
  }
  if (name === "MOD") {
    if (numbers[1] === 0) return "#DIV/0!";
    const result = numbers[0] - numbers[1] * Math.floor(numbers[0] / numbers[1]);
    return Number.isFinite(result) ? result : "#NUM!";
  }
  const result = numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
  return Number.isFinite(result) ? result : "#NUM!";
}

function formulaCondition(value) {
  if (formulaError(value)) return value;
  value = scalarFormulaValue(value);
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  const normalized = String(value).trim().toUpperCase();
  if (normalized === "TRUE") return true;
  if (normalized === "FALSE" || normalized === "") return false;
  return "#VALUE!";
}

function evaluatePivotAst(node, aggregates, options) {
  if (node.kind === "literal") return typeof node.value === "string" ? formulaText(node.value) : node.value;
  if (node.kind === "field") {
    const value = Object.hasOwn(aggregates, node.name) ? aggregates[node.name] : 0;
    if (formulaError(value)) return value;
    const number = Number(value);
    return Number.isFinite(number) ? number : 0;
  }
  if (node.kind === "unary") return formulaUnary(evaluatePivotAst(node.value, aggregates, options), (value) => node.operator === "-" ? -value : value);
  if (node.kind === "percent") return formulaUnary(evaluatePivotAst(node.value, aggregates, options), (value) => value / 100);
  if (node.kind === "binary") return formulaBinary(evaluatePivotAst(node.left, aggregates, options), node.operator, evaluatePivotAst(node.right, aggregates, options));
  if (node.name === "IF") {
    const condition = formulaCondition(evaluatePivotAst(node.args[0], aggregates, options));
    if (formulaError(condition)) return condition;
    if (condition) return evaluatePivotAst(node.args[1], aggregates, options);
    return node.args[2] ? evaluatePivotAst(node.args[2], aggregates, options) : false;
  }
  if (node.name === "IFERROR") {
    const value = evaluatePivotAst(node.args[0], aggregates, options);
    return formulaError(value) ? evaluatePivotAst(node.args[1], aggregates, options) : value;
  }
  if (node.name === "IFNA") {
    const value = evaluatePivotAst(node.args[0], aggregates, options);
    return value === "#N/A" ? evaluatePivotAst(node.args[1], aggregates, options) : value;
  }
  return formulaFunction(node.name, node.args.map((arg) => evaluatePivotAst(arg, aggregates, options)), options);
}

export function evaluatePivotFormula(formula, aggregates = {}, fields = Object.keys(aggregates), options = {}) {
  const dateSystem = options.dateSystem == null ? "1900" : String(options.dateSystem);
  if (dateSystem !== "1900" && dateSystem !== "1904") throw new TypeError("PivotTable calculated field dateSystem must be 1900 or 1904.");
  const result = evaluatePivotAst(parsedPivotFormula(formula, fields).ast, aggregates, { ...options, dateSystem });
  if (textFormulaValue(result)) return result.value;
  if (formulaError(result) || typeof result === "string" || typeof result === "boolean") return result;
  return Number.isFinite(result) ? result : "#NUM!";
}
