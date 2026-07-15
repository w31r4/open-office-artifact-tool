import { createHash } from "node:crypto";

export const PRESENTATION_OLE_WORKBOOK_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function sha256(bytes) {
  return createHash("sha256").update(bytes || []).digest("hex");
}

export async function validatePresentationOleWorkbookReplacements(objects, inspectXlsx) {
  for (const object of objects || []) {
    if (!object.oleWorkbook) continue;
    const metadata = object.oleWorkbook;
    if (object.nativeKind !== "oleObject" || metadata.contentType !== PRESENTATION_OLE_WORKBOOK_CONTENT_TYPE ||
        !/^[0-9a-f]{64}$/.test(metadata.sourceSha256) || !metadata.partPath || !metadata.relationshipId) {
      throw new Error(`Native presentation object ${object.id} has an invalid embedded-workbook source binding.`);
    }
    const matches = (object.parts || []).filter((part) => part.path === metadata.partPath && part.contentType === metadata.contentType);
    if (matches.length !== 1) throw new Error(`Native presentation object ${object.id} no longer resolves to one embedded XLSX workbook part.`);
    const part = matches[0];
    const digest = sha256(part.bytes);
    if (digest === metadata.sourceSha256) continue;
    if (!part.bytes?.length || part.bytes.length > 16 * 1024 * 1024 || part.bytes.length < 4 ||
        part.bytes[0] !== 0x50 || part.bytes[1] !== 0x4b || part.bytes[2] !== 0x03 || part.bytes[3] !== 0x04) {
      throw new Error(`Native presentation object ${object.id} replacement is not a bounded XLSX OPC package.`);
    }
    try {
      await inspectXlsx(part.bytes);
    } catch (error) {
      throw new Error(`Native presentation object ${object.id} replacement is not a valid XLSX package: ${error.message}`, { cause: error });
    }
  }
}
