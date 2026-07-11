const DEFAULT_PAGE_SIZE = { width: 612, height: 792 };

function toUint8Array(data) {
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (ArrayBuffer.isView(data)) return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  throw new Error("PDF.js parser requires binary PDF data.");
}

async function readBytes(input) {
  if (input?.bytes) return toUint8Array(input.bytes);
  if (input?.input && typeof input.input.arrayBuffer === "function") return new Uint8Array(await input.input.arrayBuffer());
  if (input?.source && typeof input.source.arrayBuffer === "function") return new Uint8Array(await input.source.arrayBuffer());
  if (input && typeof input.arrayBuffer === "function") return new Uint8Array(await input.arrayBuffer());
  return toUint8Array(input);
}

async function loadPdfjs(options = {}) {
  if (options.pdfjs?.getDocument) return options.pdfjs;
  try {
    return await import("pdfjs-dist/legacy/build/pdf.mjs");
  } catch (legacyError) {
    try {
      return await import("pdfjs-dist");
    } catch (error) {
      throw new Error(
        `PDF.js parser requires the optional peer dependency \"pdfjs-dist\". Install it with \"npm install -D pdfjs-dist\". Original error: ${error.message || legacyError.message}`,
      );
    }
  }
}

function normalizeTextItem(item, pageHeight, index) {
  const transform = item.transform || [1, 0, 0, 1, Number(item.x || 0), Number(item.y || 0)];
  const x = Number(transform[4] ?? item.x ?? 0);
  const rawY = Number(transform[5] ?? item.y ?? 0);
  const width = Number(item.width || 0);
  const height = Number(item.height || Math.abs(transform[3] || 0) || 0);
  const top = Math.max(0, pageHeight - rawY - height);
  return {
    id: `txt/${index + 1}`,
    text: String(item.str || item.text || ""),
    bbox: [x, top, width, height],
    fontName: item.fontName,
    dir: item.dir,
  };
}

function buildLines(textItems) {
  const rows = [];
  for (const item of [...textItems].sort((a, b) => (a.bbox[1] - b.bbox[1]) || (a.bbox[0] - b.bbox[0]))) {
    if (!item.text.trim()) continue;
    let row = rows.find((candidate) => Math.abs(candidate.top - item.bbox[1]) <= 4);
    if (!row) {
      row = { top: item.bbox[1], items: [] };
      rows.push(row);
    }
    row.items.push(item);
  }
  return rows.map((row) => {
    const items = row.items.sort((a, b) => a.bbox[0] - b.bbox[0]);
    return {
      text: items.map((item) => item.text).join(" ").replace(/\s+/g, " ").trim(),
      items,
      bbox: bboxForItems(items),
    };
  });
}

function bboxForItems(items) {
  if (!items.length) return [0, 0, 0, 0];
  const left = Math.min(...items.map((item) => item.bbox[0]));
  const top = Math.min(...items.map((item) => item.bbox[1]));
  const right = Math.max(...items.map((item) => item.bbox[0] + item.bbox[2]));
  const bottom = Math.max(...items.map((item) => item.bbox[1] + item.bbox[3]));
  return [left, top, Math.max(1, right - left), Math.max(1, bottom - top)];
}

function inferTables(lines, pageIndex) {
  const pipeRows = lines.filter((line) => line.text.includes("|")).map((line) => line.text.split("|").map((cell) => cell.trim()).filter(Boolean));
  if (pipeRows.length >= 2) {
    return [{ name: `pdfjs-pipe-table-${pageIndex + 1}`, values: pipeRows, bbox: bboxForItems(lines.filter((line) => line.text.includes("|")).flatMap((line) => line.items)) }];
  }

  const candidateRows = lines.map((line) => line.items.map((item) => item.text.trim()).filter(Boolean)).filter((row) => row.length >= 2);
  if (candidateRows.length < 2) return [];
  const commonWidth = candidateRows.filter((row) => Math.abs(row.length - candidateRows[0].length) <= 1);
  if (commonWidth.length < 2) return [];
  return [{ name: `pdfjs-position-table-${pageIndex + 1}`, values: commonWidth, bbox: bboxForItems(lines.flatMap((line) => line.items)) }];
}

async function extractImagePlaceholders(page, pdfjs, pageIndex, width, height) {
  const ops = pdfjs.OPS || {};
  const imageOps = new Set([ops.paintImageXObject, ops.paintJpegXObject, ops.paintInlineImageXObject, ops.paintImageMaskXObject].filter((value) => value != null));
  if (!imageOps.size || typeof page.getOperatorList !== "function") return [];
  try {
    const operatorList = await page.getOperatorList();
    let count = 0;
    return operatorList.fnArray.flatMap((fn, index) => {
      if (!imageOps.has(fn)) return [];
      count += 1;
      return [{ name: `pdfjs-image-${pageIndex + 1}-${count}`, alt: `PDF image ${count}`, bbox: [0, 0, width, height], prompt: `Extracted image operator at index ${index}` }];
    });
  } catch {
    return [];
  }
}

export async function parsePdfWithPdfjs(request = {}, defaultOptions = {}) {
  const options = { ...defaultOptions, ...(request.options?.pdfjs || {}), ...(request.pdfjsOptions || {}) };
  const pdfjs = await loadPdfjs(options);
  const bytes = await readBytes(request);
  const loadingTask = pdfjs.getDocument({
    data: bytes,
    disableFontFace: true,
    isEvalSupported: false,
    useSystemFonts: true,
    ...(options.getDocumentOptions || {}),
  });
  const document = await loadingTask.promise;
  const pages = [];

  for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber++) {
    const page = await document.getPage(pageNumber);
    const viewport = page.getViewport?.({ scale: 1 }) || DEFAULT_PAGE_SIZE;
    const width = Number(viewport.width || DEFAULT_PAGE_SIZE.width);
    const height = Number(viewport.height || DEFAULT_PAGE_SIZE.height);
    const textContent = await page.getTextContent({ normalizeWhitespace: true, disableCombineTextItems: false, ...(options.textContentOptions || {}) });
    const textItems = (textContent.items || []).map((item, index) => normalizeTextItem(item, height, index)).filter((item) => item.text);
    const lines = buildLines(textItems);
    const tables = inferTables(lines, pageNumber - 1);
    const images = await extractImagePlaceholders(page, pdfjs, pageNumber - 1, width, height);
    pages.push({
      id: `pdfjs/page/${pageNumber}`,
      width,
      height,
      text: lines.map((line) => line.text).join("\n"),
      textItems,
      regions: lines.map((line, index) => ({ id: `region/${pageNumber}/${index + 1}`, kind: "textLine", label: line.text.slice(0, 80), bbox: line.bbox })),
      tables,
      images,
    });
  }

  await loadingTask.destroy?.().catch?.(() => undefined);
  return { parser: "pdfjs", metadata: { parser: "pdfjs", pages: document.numPages }, pages };
}

export function createPdfjsParser(defaultOptions = {}) {
  return async function pdfjsParserAdapter(request = {}) {
    return parsePdfWithPdfjs(request, defaultOptions);
  };
}

export const pdfjsParser = createPdfjsParser();
