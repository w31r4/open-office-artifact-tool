import { createHash } from "node:crypto";
import { OpenChestnutCodecError } from "./open-chestnut-error.mjs";

const ASSET_PREFIX = "asset/workbook/image/";
const EMU_PER_PIXEL = 9525;
const ANCHOR_TYPES = new Set(["oneCell", "twoCell"]);
const EDIT_AS_TO_WIRE = new Map([["twoCell", 1], ["oneCell", 2], ["absolute", 3]]);
const EDIT_AS_FROM_WIRE = new Map([...EDIT_AS_TO_WIRE].map(([name, value]) => [value, name]));
const CONTENT_TYPES = new Map([
  ["image/png", { extension: "png", signature: (bytes) => bytes.length >= 8 && Buffer.from(bytes.subarray(0, 8)).equals(Buffer.from("89504e470d0a1a0a", "hex")) }],
  ["image/jpeg", { extension: "jpg", signature: (bytes) => bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff }],
]);

function fail(image, message, code = "invalid_spreadsheet_image") {
  throw new OpenChestnutCodecError(`Worksheet image ${image?.name || image?.id || "(unnamed)"} ${message}`, [], { code });
}

function assetFromImage(image) {
  if (image.uri || image.prompt) fail(image, "must use embedded dataUrl bytes; external URIs and prompts are outside the bounded OpenChestnut slice.", "unsupported_spreadsheet_image");
  if (image.fit && image.fit !== "contain") fail(image, `uses unsupported fit mode ${image.fit}; expected contain.`, "unsupported_spreadsheet_image");
  const match = /^data:(image\/(?:png|jpeg));base64,([A-Za-z0-9+/]*={0,2})$/i.exec(String(image.dataUrl || ""));
  if (!match || match[2].length === 0 || match[2].length % 4 === 1) fail(image, "requires a base64 PNG or JPEG data URL.");
  const contentType = match[1].toLowerCase();
  const profile = CONTENT_TYPES.get(contentType);
  const data = new Uint8Array(Buffer.from(match[2], "base64"));
  const canonical = Buffer.from(data).toString("base64").replace(/=+$/, "");
  if (canonical !== match[2].replace(/=+$/, "") || !profile?.signature(data)) fail(image, `data URL bytes do not match ${contentType}.`);
  const sha256 = createHash("sha256").update(data).digest("hex");
  return {
    id: `${ASSET_PREFIX}${sha256}`,
    fileName: `worksheet-image-${sha256.slice(0, 16)}.${profile.extension}`,
    contentType,
    data,
    sha256,
  };
}

function coordinate(value, name, maximum, image) {
  const number = Number(value ?? 0);
  if (!Number.isInteger(number) || number < 0 || number >= maximum) fail(image, `${name} must be an integer from 0 through ${maximum - 1}.`);
  return number;
}

function pixels(value, name, image, { positive = false } = {}) {
  const number = Number(value ?? 0);
  if (!Number.isFinite(number) || number < 0 || (positive && number <= 0) || number > 10_000_000) fail(image, `${name} must be ${positive ? "positive" : "non-negative"} and at most 10000000 pixels.`);
  return number;
}

export function spreadsheetImageSnapshot(image) {
  const anchor = image.anchor || {};
  const from = anchor.from || {};
  const to = anchor.to || {};
  const extent = anchor.extent || {};
  return {
    id: String(image.id || ""),
    name: String(image.name || ""),
    alt: String(image.alt || ""),
    dataUrl: String(image.dataUrl || ""),
    fit: String(image.fit || "contain"),
    anchorType: String(anchor.type || (anchor.to ? "twoCell" : "oneCell")),
    row: Number(from.row ?? 0),
    column: Number(from.col ?? 0),
    rowOffsetPx: Number(from.rowOffsetPx ?? 0),
    columnOffsetPx: Number(from.colOffsetPx ?? 0),
    widthPx: Number(extent.widthPx ?? image.anchor?.widthPx ?? 160),
    heightPx: Number(extent.heightPx ?? image.anchor?.heightPx ?? 120),
    toRow: anchor.to ? Number(to.row ?? 0) : undefined,
    toColumn: anchor.to ? Number(to.col ?? 0) : undefined,
    toRowOffsetPx: anchor.to ? Number(to.rowOffsetPx ?? 0) : undefined,
    toColumnOffsetPx: anchor.to ? Number(to.colOffsetPx ?? 0) : undefined,
    editAs: anchor.editAs == null ? undefined : String(anchor.editAs),
  };
}

function marker(snapshot, prefix, image) {
  const row = coordinate(snapshot[`${prefix}Row`] ?? snapshot.row, `anchor.${prefix}.row`, 1_048_576, image);
  const column = coordinate(snapshot[`${prefix}Column`] ?? snapshot.column, `anchor.${prefix}.col`, 16_384, image);
  const rowOffsetPx = pixels(snapshot[`${prefix}RowOffsetPx`] ?? snapshot.rowOffsetPx, `anchor.${prefix}.rowOffsetPx`, image);
  const columnOffsetPx = pixels(snapshot[`${prefix}ColumnOffsetPx`] ?? snapshot.columnOffsetPx, `anchor.${prefix}.colOffsetPx`, image);
  return {
    row,
    column,
    rowOffsetEmu: BigInt(Math.round(rowOffsetPx * EMU_PER_PIXEL)),
    columnOffsetEmu: BigInt(Math.round(columnOffsetPx * EMU_PER_PIXEL)),
  };
}

function markerIsAfter(to, from) {
  const columnAfter = to.column > from.column || (to.column === from.column && to.columnOffsetEmu > from.columnOffsetEmu);
  const rowAfter = to.row > from.row || (to.row === from.row && to.rowOffsetEmu > from.rowOffsetEmu);
  return columnAfter && rowAfter;
}

function wireImage(image, assets, source) {
  const snapshot = spreadsheetImageSnapshot(image);
  const asset = assetFromImage(image);
  if (!assets.has(asset.id)) assets.set(asset.id, asset);
  if (!ANCHOR_TYPES.has(snapshot.anchorType)) fail(image, `anchor.type must be oneCell or twoCell; received ${snapshot.anchorType}.`);
  const from = marker(snapshot, "from", image);
  if (!snapshot.id || snapshot.id.length > 512 || /\p{Cc}/u.test(snapshot.id)) fail(image, "id must contain 1 through 512 characters without controls.");
  if (!snapshot.name || snapshot.name.length > 255 || /\p{Cc}/u.test(snapshot.name)) fail(image, "name must contain 1 through 255 characters without controls.");
  if (snapshot.alt.length > 32_767 || /\p{Cc}/u.test(snapshot.alt)) fail(image, "alt text must contain at most 32767 characters without controls.");
  const output = {
    id: snapshot.id,
    name: snapshot.name,
    altText: snapshot.alt,
    assetId: asset.id,
    source,
  };
  if (snapshot.anchorType === "twoCell") {
    if (!image.anchor?.to) fail(image, "two-cell anchor requires anchor.to.");
    if (image.anchor?.extent || image.anchor?.widthPx != null || image.anchor?.heightPx != null) fail(image, "two-cell anchor cannot also carry one-cell extent geometry.");
    const to = marker(snapshot, "to", image);
    if (!markerIsAfter(to, from)) fail(image, "two-cell anchor.to must be strictly after anchor.from on both worksheet axes.");
    if (snapshot.editAs != null && !EDIT_AS_TO_WIRE.has(snapshot.editAs)) fail(image, `anchor.editAs must be twoCell, oneCell, or absolute; received ${snapshot.editAs}.`);
    output.twoCellAnchor = { from, to, ...(snapshot.editAs == null ? {} : { editAs: EDIT_AS_TO_WIRE.get(snapshot.editAs) }) };
  } else {
    if (image.anchor?.to) fail(image, "one-cell anchor cannot carry anchor.to.");
    if (snapshot.editAs != null) fail(image, "anchor.editAs is valid only for a two-cell anchor.");
    const widthPx = pixels(snapshot.widthPx, "anchor.extent.widthPx", image, { positive: true });
    const heightPx = pixels(snapshot.heightPx, "anchor.extent.heightPx", image, { positive: true });
    output.anchor = { ...from, widthEmu: BigInt(Math.round(widthPx * EMU_PER_PIXEL)), heightEmu: BigInt(Math.round(heightPx * EMU_PER_PIXEL)) };
  }
  return output;
}

export function wireWorksheetImages(sheet, state, assets) {
  const remaining = new Set(sheet.images?.items || []);
  const output = [];
  for (const slot of state?.slots || []) {
    if (!remaining.delete(slot.image)) {
      throw new OpenChestnutCodecError(`Worksheet ${sheet.name} cannot remove imported image ${slot.image?.name || slot.wire.id} in the bounded OpenChestnut slice.`, [], { code: "invalid_spreadsheet_image_topology" });
    }
    const unchanged = JSON.stringify(spreadsheetImageSnapshot(slot.image)) === JSON.stringify(slot.publicSnapshot);
    if (unchanged) {
      output.push(slot.wire);
      const asset = assetFromImage(slot.image);
      if (!assets.has(asset.id)) assets.set(asset.id, asset);
    } else output.push(wireImage(slot.image, assets, slot.wire.source));
  }
  if (state && remaining.size) {
    const image = [...remaining][0];
    throw new OpenChestnutCodecError(`Worksheet ${sheet.name} cannot add image ${image.name} to an imported source package in the bounded OpenChestnut slice.`, [], { code: "invalid_spreadsheet_image_topology" });
  }
  output.push(...[...remaining].map((image) => wireImage(image, assets)));
  return output;
}

function assetBytes(value) {
  if (value instanceof Uint8Array) return value;
  if (ArrayBuffer.isView(value)) return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  throw new TypeError("Spreadsheet image asset data must be bytes.");
}

function dataUrl(asset, image) {
  if (!asset) fail(image, `references missing asset ${image.assetId}.`);
  const profile = CONTENT_TYPES.get(String(asset.contentType || "").toLowerCase());
  const bytes = assetBytes(asset.data);
  if (!profile?.signature(bytes)) fail(image, `asset ${asset.id} bytes do not match a supported PNG or JPEG content type.`);
  const digest = createHash("sha256").update(bytes).digest("hex");
  if (asset.sha256 !== digest || asset.id !== `${ASSET_PREFIX}${digest}`) fail(image, `asset ${asset.id} is not content-addressed by its bytes.`);
  return `data:${asset.contentType};base64,${Buffer.from(bytes).toString("base64")}`;
}

export function spreadsheetImageFromWire(sheet, source, assets) {
  const anchor = source.anchor;
  const twoCellAnchor = source.twoCellAnchor;
  if (Boolean(anchor) === Boolean(twoCellAnchor)) fail(source, "must carry exactly one one-cell or two-cell anchor.");
  const numberFromWire = (value, name) => {
    const number = Number(value);
    if (!Number.isSafeInteger(number) || number < 0) fail(source, `has invalid ${name}.`);
    return number;
  };
  const publicMarker = (value, name) => {
    if (!value) fail(source, `has no ${name} marker.`);
    const row = Number(value.row);
    const col = Number(value.column);
    if (!Number.isInteger(row) || row < 0 || row >= 1_048_576 || !Number.isInteger(col) || col < 0 || col >= 16_384) fail(source, `has invalid ${name} row/column coordinates.`);
    return {
      row,
      col,
      rowOffsetPx: numberFromWire(value.rowOffsetEmu, `${name} row offset`) / EMU_PER_PIXEL,
      colOffsetPx: numberFromWire(value.columnOffsetEmu, `${name} column offset`) / EMU_PER_PIXEL,
    };
  };
  let publicAnchor;
  if (twoCellAnchor) {
    const editAs = twoCellAnchor.editAs == null ? undefined : EDIT_AS_FROM_WIRE.get(twoCellAnchor.editAs);
    if (twoCellAnchor.editAs != null && !editAs) fail(source, `has invalid two-cell editAs value ${twoCellAnchor.editAs}.`);
    publicAnchor = {
      type: "twoCell",
      from: publicMarker(twoCellAnchor.from, "from"),
      to: publicMarker(twoCellAnchor.to, "to"),
      ...(editAs ? { editAs } : {}),
    };
  } else {
    publicAnchor = {
      from: publicMarker(anchor, "from"),
      extent: {
        widthPx: numberFromWire(anchor.widthEmu, "width") / EMU_PER_PIXEL,
        heightPx: numberFromWire(anchor.heightEmu, "height") / EMU_PER_PIXEL,
      },
    };
  }
  const image = sheet.images.add({
    id: source.id,
    name: source.name,
    alt: source.altText,
    dataUrl: dataUrl(assets.get(source.assetId), source),
    fit: "contain",
    anchor: publicAnchor,
  });
  image.id = source.id || image.id;
  return image;
}
