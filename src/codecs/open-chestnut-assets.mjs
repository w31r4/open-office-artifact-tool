import { createHash } from "node:crypto";
import { OpenChestnutCodecError } from "./open-chestnut-error.mjs";

const MAX_ASSET_BYTES = 16 * 1024 * 1024;
const MAX_ASSET_COUNT = 1024;
const PICTURE_ASSET_PREFIX = "asset/presentation/picture-bullet/";
const OLE_WORKBOOK_ASSET_PREFIX = "asset/presentation/ole-workbook/";
const XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const IMAGE_TYPES = new Map([
  ["image/png", { extension: "png", signature: (bytes) => bytes.length >= 8 && bytes.subarray(0, 8).equals(Buffer.from("89504e470d0a1a0a", "hex")) }],
  ["image/jpeg", { extension: "jpg", signature: (bytes) => bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff }],
  ["image/gif", { extension: "gif", signature: (bytes) => bytes.length >= 6 && new Set(["GIF87a", "GIF89a"]).has(bytes.subarray(0, 6).toString("ascii")) }],
  ["image/svg+xml", { extension: "svg", signature: safeSvg }],
]);

function fail(message, code = "invalid_presentation_asset") {
  throw new OpenChestnutCodecError(message, [], { code });
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function safeSvg(bytes) {
  const source = bytes.toString("utf8").replace(/^\uFEFF/, "");
  if (!/^\s*(?:<\?xml[\s\S]*?\?>\s*)?<svg(?:\s|>)/i.test(source)) return false;
  if (/<!DOCTYPE|<!ENTITY|<\?xml-stylesheet|<\s*(?:script|foreignObject)\b|\son[a-z]+\s*=|@import\b/i.test(source)) return false;
  for (const match of source.matchAll(/\s(?:href|xlink:href)\s*=\s*(["'])(.*?)\1/gi)) {
    const target = match[2].trim();
    if (target && !target.startsWith("#") && !/^data:image\/(?:png|jpe?g|gif);base64,/i.test(target)) return false;
  }
  for (const match of source.matchAll(/url\(\s*(["']?)(.*?)\1\s*\)/gi)) {
    const target = match[2].trim();
    if (target && !target.startsWith("#") && !/^data:image\/(?:png|jpe?g|gif);base64,/i.test(target)) return false;
  }
  return true;
}

function normalizeContentType(value) {
  const contentType = String(value || "").toLowerCase();
  if (contentType === "image/jpg") return "image/jpeg";
  return contentType;
}

function validateImage(contentType, bytes, label) {
  const format = IMAGE_TYPES.get(contentType);
  if (!format) fail(`${label} must be a PNG, JPEG, GIF, or safe SVG image.`);
  if (!bytes.length || bytes.length > MAX_ASSET_BYTES) fail(`${label} must contain 1 through ${MAX_ASSET_BYTES} bytes.`);
  if (!format.signature(bytes)) fail(`${label} bytes do not match content type ${contentType}.`);
  return format;
}

function parseDataUrl(value) {
  const source = String(value);
  if (source.length > Math.ceil(MAX_ASSET_BYTES / 3) * 4 + 128) fail(`Presentation picture bullet must not exceed ${MAX_ASSET_BYTES} decoded bytes.`, "presentation_asset_budget_exceeded");
  const match = /^data:(image\/(?:png|jpe?g|gif|svg\+xml));base64,([A-Za-z0-9+/=\s]+)$/i.exec(source);
  if (!match) fail("Presentation picture bullet dataUrl must be a base64 PNG, JPEG, GIF, or SVG image.");
  const contentType = normalizeContentType(match[1]);
  const encoded = match[2].replace(/\s/g, "");
  if (encoded.length % 4 === 1 || /=[^=]|={3,}/.test(encoded)) fail("Presentation picture bullet dataUrl contains invalid base64.");
  const bytes = Buffer.from(encoded, "base64");
  const canonical = bytes.toString("base64").replace(/=+$/, "");
  if (canonical !== encoded.replace(/=+$/, "")) fail("Presentation picture bullet dataUrl contains invalid base64.");
  const format = validateImage(contentType, bytes, "Presentation picture bullet");
  return { bytes, contentType, extension: format.extension };
}

function validateAsset(asset) {
  const id = String(asset?.id || "");
  const isPicture = new RegExp(`^${PICTURE_ASSET_PREFIX}[0-9a-f]{64}$`).test(id);
  const isOleWorkbook = new RegExp(`^${OLE_WORKBOOK_ASSET_PREFIX}[0-9a-f]{64}$`).test(id);
  if (!isPicture && !isOleWorkbook) fail(`Presentation asset ID ${id || "(missing)"} is invalid.`);
  const contentType = normalizeContentType(asset.contentType);
  const bytes = Buffer.from(asset.data || []);
  const format = isPicture
    ? validateImage(contentType, bytes, `Presentation asset ${id}`)
    : validateOleWorkbook(contentType, bytes, `Presentation asset ${id}`);
  const digest = sha256(bytes);
  if (String(asset.sha256 || "").toLowerCase() !== digest) fail(`Presentation asset ${id} does not match its SHA-256 digest.`);
  const expectedId = `${isPicture ? PICTURE_ASSET_PREFIX : OLE_WORKBOOK_ASSET_PREFIX}${digest}`;
  if (id !== expectedId) fail(`Presentation asset ${id} is not content-addressed by its bytes.`);
  return {
    id,
    fileName: String(asset.fileName || (isPicture ? `picture-bullet-${digest.slice(0, 16)}.${format.extension}` : `embedded-workbook-${digest.slice(0, 16)}.xlsx`)),
    contentType,
    data: bytes,
    sha256: digest,
  };
}

function validateOleWorkbook(contentType, bytes, label) {
  if (contentType !== XLSX_CONTENT_TYPE) fail(`${label} must use the XLSX workbook content type.`);
  if (!bytes.length || bytes.length > MAX_ASSET_BYTES) fail(`${label} must contain 1 through ${MAX_ASSET_BYTES} bytes.`);
  if (bytes.length < 4 || !bytes.subarray(0, 4).equals(Buffer.from("504b0304", "hex"))) fail(`${label} must contain an OPC ZIP package.`);
  return { extension: "xlsx" };
}

export function createPresentationAssetCatalog(initialAssets = []) {
  if (initialAssets.length > MAX_ASSET_COUNT) fail(`Presentation exceeds the ${MAX_ASSET_COUNT}-asset budget.`, "presentation_asset_budget_exceeded");
  const byId = new Map();
  for (const raw of initialAssets) {
    const asset = validateAsset(raw);
    if (byId.has(asset.id)) fail(`Presentation contains duplicate asset ID ${asset.id}.`);
    byId.set(asset.id, asset);
  }
  return {
    addDataUrl(dataUrl) {
      const decoded = parseDataUrl(dataUrl);
      const digest = sha256(decoded.bytes);
      const id = `${PICTURE_ASSET_PREFIX}${digest}`;
      if (!byId.has(id)) {
        if (byId.size >= MAX_ASSET_COUNT) fail(`Presentation exceeds the ${MAX_ASSET_COUNT}-asset budget.`, "presentation_asset_budget_exceeded");
        byId.set(id, {
          id,
          fileName: `picture-bullet-${digest.slice(0, 16)}.${decoded.extension}`,
          contentType: decoded.contentType,
          data: decoded.bytes,
          sha256: digest,
        });
      }
      return id;
    },
    dataUrl(id) {
      const asset = byId.get(String(id));
      if (!asset || !asset.id.startsWith(PICTURE_ASSET_PREFIX)) fail(`Presentation picture bullet references missing asset ${id || "(missing)"}.`);
      return `data:${asset.contentType};base64,${asset.data.toString("base64")}`;
    },
    assets() {
      return [...byId.values()].map((asset) => ({ ...asset, data: Uint8Array.from(asset.data) }));
    },
  };
}

export function validatePictureBulletUri(value) {
  const uri = String(value || "");
  if (!uri || uri.length > 4096 || /[\u0000-\u001f\u007f]/.test(uri)) fail("Presentation picture bullet URI must contain 1 through 4096 characters without controls.");
  let parsed;
  try {
    parsed = new URL(uri);
  } catch {
    fail("Presentation picture bullet URI must be absolute.");
  }
  if (!new Set(["http:", "https:"]).has(parsed.protocol)) fail("Presentation picture bullet URI must use http or https.");
  return uri;
}
