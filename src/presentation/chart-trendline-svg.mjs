import { resolveColorToken } from "../shared/colors.mjs";

const CURVE_SAMPLE_COUNT = 65;
const SOLVER_EPSILON = 1e-12;

function attrEscape(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function svgNumber(value) {
  const rounded = Math.round(value * 1_000) / 1_000;
  return Object.is(rounded, -0) ? "0" : String(rounded);
}

export function presentationChartLineSvgAttributes(line) {
  if (!line) return "";
  const dash = {
    dot: "1 3",
    dash: "6 4",
    longDash: "10 4",
    dashDot: "6 3 1 3",
    longDashDot: "10 4 1 4",
    longDashDotDot: "10 3 1 3 1 3",
    systemDash: "4 3",
    systemDot: "1 2",
    systemDashDot: "4 2 1 2",
    systemDashDotDot: "4 2 1 2 1 2",
  }[line.style];
  return ` stroke="${attrEscape(resolveColorToken(line.fill, line.fill || "#0f172a"))}" stroke-width="${svgNumber(line.width)}"${dash ? ` stroke-dasharray="${dash}"` : ""}`;
}

function finiteSeriesPoints(values = []) {
  return values.map((value, index) => ({ x: index + 1, y: Number(value) }))
    .filter((point) => Number.isFinite(point.y));
}

function solveLinearSystem(matrix, vector) {
  const size = vector.length;
  const augmented = matrix.map((row, index) => [...row, vector[index]]);
  for (let column = 0; column < size; column++) {
    let pivot = column;
    for (let row = column + 1; row < size; row++) {
      if (Math.abs(augmented[row][column]) > Math.abs(augmented[pivot][column])) pivot = row;
    }
    if (Math.abs(augmented[pivot][column]) <= SOLVER_EPSILON) return undefined;
    [augmented[column], augmented[pivot]] = [augmented[pivot], augmented[column]];
    const divisor = augmented[column][column];
    for (let index = column; index <= size; index++) augmented[column][index] /= divisor;
    for (let row = 0; row < size; row++) {
      if (row === column) continue;
      const factor = augmented[row][column];
      for (let index = column; index <= size; index++) augmented[row][index] -= factor * augmented[column][index];
    }
  }
  const result = augmented.map((row) => row[size]);
  return result.every(Number.isFinite) ? result : undefined;
}

function leastSquares(points, basisFunctions) {
  if (points.length < basisFunctions.length || basisFunctions.length === 0) return undefined;
  const matrix = basisFunctions.map((left) => basisFunctions.map((right) => points.reduce((sum, point) => sum + left(point.x) * right(point.x), 0)));
  const vector = basisFunctions.map((basis) => points.reduce((sum, point) => sum + basis(point.x) * point.y, 0));
  return solveLinearSystem(matrix, vector);
}

function polynomialPredictor(points, order, fixedIntercept) {
  const scale = Math.max(1, ...points.map((point) => Math.abs(point.x)));
  const firstPower = fixedIntercept == null ? 0 : 1;
  const powers = Array.from({ length: order - firstPower + 1 }, (_, index) => index + firstPower);
  const shifted = points.map((point) => ({ x: point.x / scale, y: point.y - (fixedIntercept ?? 0) }));
  const coefficients = leastSquares(shifted, powers.map((power) => (x) => x ** power));
  if (!coefficients) return undefined;
  return (x) => (fixedIntercept ?? 0) + coefficients.reduce((sum, coefficient, index) => sum + coefficient * (x / scale) ** powers[index], 0);
}

function transformedLinearPredictor(points, transformX, transformY, restoreY, transformedIntercept) {
  const transformed = [];
  for (const point of points) {
    const x = transformX(point.x);
    const y = transformY(point.y);
    if (Number.isFinite(x) && Number.isFinite(y)) transformed.push({ x, y });
  }
  if (transformed.length !== points.length || transformed.length < 2) return undefined;
  const shifted = transformedIntercept == null
    ? transformed
    : transformed.map((point) => ({ ...point, y: point.y - transformedIntercept }));
  const basis = transformedIntercept == null ? [() => 1, (x) => x] : [(x) => x];
  const coefficients = leastSquares(shifted, basis);
  if (!coefficients) return undefined;
  return (rawX) => {
    const x = transformX(rawX);
    if (!Number.isFinite(x)) return Number.NaN;
    const fitted = (transformedIntercept ?? 0) + coefficients.reduce((sum, coefficient, index) => sum + coefficient * basis[index](x), 0);
    return restoreY(fitted);
  };
}

function trendlinePredictor(trendline, points) {
  if (trendline.type === "linear") return polynomialPredictor(points, 1, trendline.intercept);
  if (trendline.type === "poly") return polynomialPredictor(points, trendline.order || 2, trendline.intercept);
  if (trendline.type === "exp") {
    if (trendline.intercept != null && !(trendline.intercept > 0)) return undefined;
    return transformedLinearPredictor(points, (x) => x, (y) => y > 0 ? Math.log(y) : Number.NaN, Math.exp, trendline.intercept == null ? undefined : Math.log(trendline.intercept));
  }
  if (trendline.type === "log") return transformedLinearPredictor(points, (x) => x > 0 ? Math.log(x) : Number.NaN, (y) => y, (y) => y, trendline.intercept);
  if (trendline.type === "power") {
    if (trendline.intercept != null && !(trendline.intercept > 0)) return undefined;
    return transformedLinearPredictor(points, (x) => x > 0 ? Math.log(x) : Number.NaN, (y) => y > 0 ? Math.log(y) : Number.NaN, Math.exp, trendline.intercept == null ? undefined : Math.log(trendline.intercept));
  }
  return undefined;
}

function movingAveragePoints(values, period) {
  const points = [];
  for (let end = period - 1; end < values.length; end++) {
    const window = values.slice(end - period + 1, end + 1).map(Number);
    if (!window.every(Number.isFinite)) continue;
    points.push({ x: end + 1, y: window.reduce((sum, value) => sum + value, 0) / period });
  }
  return points;
}

function trendlineDomain(trendline, categoryCount) {
  if (trendline.type === "movingAvg") return { start: 1, end: categoryCount };
  return {
    start: 1 - (trendline.backward || 0),
    end: categoryCount + (trendline.forward || 0),
  };
}

function mapPoint(point, plot, max, domain, { horizontal, centered }) {
  const axisDomain = centered ? { start: domain.start - 0.5, end: domain.end + 0.5 } : domain;
  const domainWidth = axisDomain.end - axisDomain.start;
  if (!(domainWidth > 0) || !Number.isFinite(point.x) || !Number.isFinite(point.y) || point.y < 0 || point.y > max) return undefined;
  const categoryRatio = (point.x - axisDomain.start) / domainWidth;
  const valueRatio = point.y / max;
  if (horizontal) {
    return {
      x: plot.left + valueRatio * plot.width,
      y: plot.top + categoryRatio * plot.height,
    };
  }
  return {
    x: plot.left + categoryRatio * plot.width,
    y: plot.top + plot.height - valueRatio * plot.height,
  };
}

function curveSegments(predict, domain, sampleCount = CURVE_SAMPLE_COUNT) {
  const segments = [];
  let current = [];
  const count = Math.max(2, Math.min(257, Math.trunc(sampleCount) || CURVE_SAMPLE_COUNT));
  for (let index = 0; index < count; index++) {
    const x = domain.start + (index / (count - 1)) * (domain.end - domain.start);
    const point = { x, y: predict(x) };
    if (Number.isFinite(point.y)) current.push(point);
    else if (current.length) {
      if (current.length > 1) segments.push(current);
      current = [];
    }
  }
  if (current.length > 1) segments.push(current);
  return segments;
}

export function samplePresentationChartTrendline(values, trendline, options = {}) {
  const categoryCount = Math.max(0, Math.trunc(options.categoryCount ?? values?.length ?? 0));
  if (categoryCount < 2) return [];
  const domain = trendlineDomain(trendline, categoryCount);
  if (!(domain.end > domain.start)) return [];
  if (trendline.type === "movingAvg") {
    const points = movingAveragePoints(values || [], trendline.period || 2);
    return points.length > 1 ? [{ domain, points }] : [];
  }
  const sourcePoints = finiteSeriesPoints(values);
  if (sourcePoints.length < 2) return [];
  const predict = trendlinePredictor(trendline, sourcePoints);
  return predict ? curveSegments(predict, domain, options.sampleCount).map((points) => ({ domain, points })) : [];
}

function polylineSvg(points, trendline, attributes) {
  if (points.length < 2) return "";
  const encoded = points.map((point) => `${svgNumber(point.x)},${svgNumber(point.y)}`).join(" ");
  return `<polyline data-trendline-type="${trendline.type}" points="${encoded}" fill="none"${attributes}/>`;
}

export function presentationChartTrendlinesSvg(series, plot, max, categoryCount, { horizontal = false, centered = horizontal } = {}) {
  if (categoryCount < 2 || !(max > 0)) return "";
  return (series.trendlines || []).map((trendline) => {
    const line = trendline.line || { fill: series.color || "#475569", width: 1.5, style: "dash" };
    const attributes = presentationChartLineSvgAttributes(line);
    return samplePresentationChartTrendline(series.values || [], trendline, { categoryCount })
      .flatMap(({ domain, points }) => {
        const segments = [];
        let current = [];
        for (const point of points) {
          const mapped = mapPoint(point, plot, max, domain, { horizontal, centered });
          if (mapped) current.push(mapped);
          else if (current.length) {
            if (current.length > 1) segments.push(current);
            current = [];
          }
        }
        if (current.length > 1) segments.push(current);
        return segments;
      })
      .map((points) => polylineSvg(points, trendline, attributes))
      .join("");
  }).join("");
}
