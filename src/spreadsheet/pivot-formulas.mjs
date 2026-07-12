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
    if ("+-*/^()%".includes(text[index])) { tokens.push({ type: "operator", value: text[index++] }); continue; }
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
  return tokens.map((token) => token.type === "field" ? `'${token.value.replaceAll("'", "''")}'` : String(token.value)).join("");
}

function uniqueReferences(tokens) {
  return [...new Set(tokens.filter((token) => token.type === "field").map((token) => token.value))];
}

export function pivotFormulaToOoxml(formula, fields) {
  return renderPivotFormula(formulaTokens(formula, fields));
}

export function normalizeCalculatedFields(value, sourceFields, allowUnsupported = false) {
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
      const tokens = formulaTokens(formula, sourceFields);
      evaluatePivotFormula(formula, Object.fromEntries(sourceFields.map((field) => [field, 0])), sourceFields);
      return { name, formula: `=${renderPivotFormula(tokens)}`, numFmtId, references: uniqueReferences(tokens) };
    } catch (error) {
      if (!allowUnsupported || error instanceof RangeError) throw error;
      return { name, formula: formula.startsWith("=") ? formula : `=${formula}`, numFmtId, references: [], supported: false, error: error.message };
    }
  });
}

function formulaBinary(left, operator, right) {
  if (typeof left === "string" && left.startsWith("#")) return left;
  if (typeof right === "string" && right.startsWith("#")) return right;
  if (operator === "/" && right === 0) return "#DIV/0!";
  if (operator === "+") return left + right;
  if (operator === "-") return left - right;
  if (operator === "*") return left * right;
  if (operator === "/") return left / right;
  return left ** right;
}

function formulaUnary(value, transform) {
  return typeof value === "string" && value.startsWith("#") ? value : transform(Number(value));
}

export function evaluatePivotFormula(formula, aggregates = {}, fields = Object.keys(aggregates)) {
  const tokens = formulaTokens(formula, fields);
  let index = 0;
  const primary = () => {
    const token = tokens[index++];
    if (!token) throw new SyntaxError("PivotTable calculated field formula ended unexpectedly.");
    if (token.type === "number") return token.value;
    if (token.type === "field") return Number(aggregates[token.value]) || 0;
    if (token.value === "(") { const value = expression(); if (tokens[index++]?.value !== ")") throw new SyntaxError("PivotTable calculated field formula requires a closing parenthesis."); return value; }
    throw new SyntaxError(`PivotTable calculated field formula has unexpected token ${token.value}.`);
  };
  const unary = () => tokens[index]?.value === "+" ? (index++, formulaUnary(unary(), (value) => value)) : tokens[index]?.value === "-" ? (index++, formulaUnary(unary(), (value) => -value)) : primary();
  const power = () => { let left = unary(); if (tokens[index]?.value === "^") { index += 1; left = formulaBinary(left, "^", power()); } return left; };
  const percent = () => { let value = power(); while (tokens[index]?.value === "%") { index += 1; value = formulaUnary(value, (number) => number / 100); } return value; };
  const term = () => { let left = percent(); while (["*", "/"].includes(tokens[index]?.value)) { const operator = tokens[index++].value; left = formulaBinary(left, operator, percent()); } return left; };
  const expression = () => { let left = term(); while (["+", "-"].includes(tokens[index]?.value)) { const operator = tokens[index++].value; left = formulaBinary(left, operator, term()); } return left; };
  const result = expression();
  if (index !== tokens.length) throw new SyntaxError(`PivotTable calculated field formula has unexpected token ${tokens[index].value}.`);
  return Number.isFinite(result) || (typeof result === "string" && result.startsWith("#")) ? result : "#NUM!";
}
