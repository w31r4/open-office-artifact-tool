const COLOR_TOKENS = Object.freeze({
  white: "#ffffff",
  black: "#000000",
  transparent: "transparent",
  "slate-50": "#f8fafc",
  "slate-100": "#f1f5f9",
  "slate-200": "#e2e8f0",
  "slate-300": "#cbd5e1",
  "slate-500": "#64748b",
  "slate-600": "#475569",
  "slate-700": "#334155",
  "slate-900": "#0f172a",
  "slate-950": "#020617",
  "sky-50": "#f0f9ff",
  "sky-100": "#e0f2fe",
  "sky-500": "#0ea5e9",
  "sky-600": "#0284c7",
  green: "#22c55e",
  red: "#ef4444",
  accent1: "#156082",
  tx1: "#1f1f1f",
  bg1: "#ffffff",
});

export function resolveColorToken(value, fallback = "transparent") {
  if (!value) return fallback;
  const raw = String(value);
  return COLOR_TOKENS[raw.split("/")[0]] || raw;
}
