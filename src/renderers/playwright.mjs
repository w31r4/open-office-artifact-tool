import { FileBlob } from "../index.mjs";

const MIME_BY_FORMAT = {
  png: "image/png",
  webp: "image/webp",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  pdf: "application/pdf",
};

const DEFAULT_VIEWPORT = { width: 1280, height: 720 };

function normalizeMime(type = "") {
  return String(type || "").split(";")[0].trim().toLowerCase();
}

function normalizeFormat(format, outputType) {
  const raw = String(format || "").trim().toLowerCase();
  if (raw) {
    if (raw === "image/png") return "png";
    if (raw === "image/webp") return "webp";
    if (raw === "image/jpeg") return "jpeg";
    if (raw === "application/pdf") return "pdf";
    return raw;
  }
  const type = normalizeMime(outputType);
  return Object.entries(MIME_BY_FORMAT).find(([, mime]) => mime === type)?.[0] || "png";
}

function parsePositiveNumber(value) {
  if (value == null || value === "") return undefined;
  const text = String(value).trim();
  if (text.endsWith("%")) return undefined;
  const number = Number.parseFloat(text);
  return Number.isFinite(number) && number > 0 ? number : undefined;
}

function inferSvgViewport(source) {
  const svgOpen = String(source || "").match(/<svg\b[^>]*>/i)?.[0] || "";
  const width = parsePositiveNumber(svgOpen.match(/\bwidth=["']([^"']+)["']/i)?.[1]);
  const height = parsePositiveNumber(svgOpen.match(/\bheight=["']([^"']+)["']/i)?.[1]);
  if (width && height) return { width: Math.ceil(width), height: Math.ceil(height) };

  const viewBox = svgOpen.match(/\bviewBox=["']([^"']+)["']/i)?.[1];
  if (viewBox) {
    const parts = viewBox.trim().split(/[\s,]+/).map(Number);
    if (parts.length === 4 && parts.every(Number.isFinite) && parts[2] > 0 && parts[3] > 0) {
      return { width: Math.ceil(parts[2]), height: Math.ceil(parts[3]) };
    }
  }
  return null;
}

function normalizeViewport(optionViewport, source, inputType) {
  const inferred = inputType === "image/svg+xml" ? inferSvgViewport(source) : null;
  const viewport = optionViewport || inferred || DEFAULT_VIEWPORT;
  return {
    width: Math.max(1, Math.round(parsePositiveNumber(viewport.width) || inferred?.width || DEFAULT_VIEWPORT.width)),
    height: Math.max(1, Math.round(parsePositiveNumber(viewport.height) || inferred?.height || DEFAULT_VIEWPORT.height)),
  };
}

function deterministicCss({ background = "white", viewport }) {
  return `
    html, body {
      margin: 0;
      width: ${viewport.width}px;
      min-width: ${viewport.width}px;
      height: ${viewport.height}px;
      min-height: ${viewport.height}px;
      overflow: hidden;
      background: ${background};
    }
    *, *::before, *::after {
      animation-delay: 0s !important;
      animation-duration: 0s !important;
      animation-iteration-count: 1 !important;
      caret-color: transparent !important;
      transition-delay: 0s !important;
      transition-duration: 0s !important;
    }
    body > svg:first-child {
      display: block;
      max-width: 100%;
      max-height: 100%;
    }
  `;
}

function cspMeta(allowNetwork) {
  if (allowNetwork) return "";
  return `<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data: blob:; style-src 'unsafe-inline'; font-src data:; script-src 'none';">`;
}

function wrapHtml(source, inputType, renderOptions) {
  const style = `<style>${deterministicCss(renderOptions)}</style>`;
  const csp = cspMeta(renderOptions.allowNetwork);
  if (inputType === "image/svg+xml") {
    return `<!doctype html><html><head><meta charset="utf-8">${csp}${style}</head><body>${source}</body></html>`;
  }

  const html = String(source || "");
  if (/<head\b[^>]*>/i.test(html)) {
    return html.replace(/<head\b[^>]*>/i, (match) => `${match}<meta charset="utf-8">${csp}${style}`);
  }
  if (/<html\b[^>]*>/i.test(html)) {
    return html.replace(/<html\b[^>]*>/i, (match) => `${match}<head><meta charset="utf-8">${csp}${style}</head>`);
  }
  return `<!doctype html><html><head><meta charset="utf-8">${csp}${style}</head><body>${html}</body></html>`;
}

function mergeAdapterOptions(defaultOptions = {}, request = {}) {
  return {
    ...defaultOptions,
    ...(request.options?.playwright || {}),
    ...(request.playwright || {}),
  };
}

async function readInputText(input) {
  if (typeof input === "string") return input;
  if (input && typeof input.text === "function") return input.text();
  if (input && typeof input.arrayBuffer === "function") return new TextDecoder().decode(new Uint8Array(await input.arrayBuffer()));
  throw new Error("Playwright renderer requires an input FileBlob, Blob, or string.");
}

async function loadPlaywright(options = {}) {
  if (options.chromium) return { chromium: options.chromium };
  if (options.playwright?.chromium) return options.playwright;
  try {
    return await import("playwright");
  } catch (error) {
    throw new Error(
      `Playwright renderer requires the optional peer dependency \"playwright\". Install it with \"npm install -D playwright\" and install Chromium with \"npx playwright install chromium\". Original error: ${error.message}`,
    );
  }
}

function shouldBlockUrl(url) {
  return !/^(about:blank|data:|blob:)/i.test(String(url || ""));
}

function metadataFor(request, options, viewport, outputType, inputType, format) {
  return {
    renderer: "playwright",
    artifactKind: request.artifactKind,
    format,
    inputType,
    outputType,
    viewport,
    deviceScaleFactor: options.deviceScaleFactor ?? 1,
    network: options.allowNetwork ? "allowed" : "blocked",
  };
}

export async function renderWithPlaywright(request = {}, defaultOptions = {}) {
  const input = request.input || request.source;
  const inputType = normalizeMime(request.inputType || input?.type || defaultOptions.inputType || "image/svg+xml");
  if (!input) throw new Error("Playwright renderer requires request.input or request.source.");
  if (!["image/svg+xml", "text/html", "application/xhtml+xml"].includes(inputType)) {
    throw new Error(`Playwright renderer supports SVG or HTML input, not ${inputType || "unknown"}.`);
  }

  const options = mergeAdapterOptions(defaultOptions, request);
  const source = await readInputText(input);
  const format = normalizeFormat(request.format || options.format, request.outputType);
  const outputType = request.outputType || MIME_BY_FORMAT[format];
  if (!outputType || !MIME_BY_FORMAT[format]) {
    throw new Error(`Playwright renderer cannot produce ${request.format || request.outputType || "unknown"}; supported formats are png, webp, jpeg, and pdf.`);
  }

  const viewport = normalizeViewport(options.viewport, source, inputType);
  const renderOptions = {
    allowNetwork: options.allowNetwork === true,
    background: options.background || "white",
    viewport,
  };
  const html = wrapHtml(source, inputType, renderOptions);
  const { chromium } = await loadPlaywright(options);
  const timeout = options.timeout ?? 30_000;
  const deviceScaleFactor = options.deviceScaleFactor ?? 1;
  const ownsBrowser = !options.browser;
  const browser = options.browser || await chromium.launch({ headless: true, ...(options.launchOptions || {}) });
  let context;

  try {
    context = await browser.newContext({
      viewport,
      deviceScaleFactor,
      colorScheme: "light",
      reducedMotion: "reduce",
      locale: "en-US",
      timezoneId: "UTC",
      serviceWorkers: "block",
      ...(options.contextOptions || {}),
      viewport,
      deviceScaleFactor,
    });
    if (!renderOptions.allowNetwork) {
      await context.route("**/*", (route) => {
        if (shouldBlockUrl(route.request().url())) return route.abort("blockedbyclient");
        return route.continue();
      });
    }

    const page = await context.newPage();
    page.setDefaultTimeout(timeout);
    await page.setContent(html, { waitUntil: "domcontentloaded", timeout });
    await page.evaluate(() => document.fonts?.ready).catch(() => undefined);
    await page.evaluate(() => new Promise((resolve) => requestAnimationFrame(() => resolve()))).catch(() => undefined);

    let bytes;
    if (format === "pdf") {
      bytes = await page.pdf({
        width: `${viewport.width}px`,
        height: `${viewport.height}px`,
        printBackground: true,
        preferCSSPageSize: false,
        ...(options.pdfOptions || {}),
      });
    } else {
      const screenshotType = format === "jpg" ? "jpeg" : format;
      bytes = await page.screenshot({
        type: screenshotType,
        fullPage: false,
        omitBackground: options.omitBackground ?? false,
        animations: "disabled",
        ...(options.screenshotOptions || {}),
      });
    }

    return new FileBlob(bytes, { type: outputType, metadata: metadataFor(request, options, viewport, outputType, inputType, format) });
  } finally {
    await context?.close().catch(() => undefined);
    if (ownsBrowser) await browser.close().catch(() => undefined);
  }
}

export function createPlaywrightRenderer(defaultOptions = {}) {
  return async function playwrightRendererAdapter(request = {}) {
    return renderWithPlaywright(request, defaultOptions);
  };
}

export const playwrightRenderer = createPlaywrightRenderer();
