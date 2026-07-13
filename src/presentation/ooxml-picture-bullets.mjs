import path from "node:path";
import { Buffer } from "node:buffer";

export const PRESENTATION_IMAGE_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

function pictureBulletSource(bulletImage) {
  return bulletImage?.dataUrl || bulletImage?.uri;
}

export function presentationPictureBulletReferencesFromParagraphs(paragraphs = []) {
  return paragraphs.map((paragraph) => paragraph?.bulletImage).filter(Boolean);
}

export function presentationPictureBulletReferencesFromStyles(styles = {}) {
  return Object.values(styles || {}).map((style) => style?.bulletImage).filter(Boolean);
}

export function planPresentationPictureBullets(options = {}) {
  const owners = options.owners || [];
  const decodeDataUrl = options.decodeDataUrl;
  const existingMediaParts = options.existingMediaParts || [];
  const mediaBySource = new Map(existingMediaParts.filter((part) => part.source).map((part) => [part.source, part]));
  const mediaParts = [];
  let nextMediaPartId = Math.max(0, ...existingMediaParts.map((part) => Number(part.imagePartId) || 0)) + 1;
  const byOwner = new Map();

  for (const owner of owners) {
    let nextRelationshipIndex = Number(owner.startRelationshipIndex || 1);
    if (!Number.isInteger(nextRelationshipIndex) || nextRelationshipIndex < 1) throw new RangeError("Presentation picture-bullet relationship index must be a positive integer.");
    const relationshipIds = new Map((owner.existingRelationships || []).filter((relationship) => relationship.source).map((relationship) => [relationship.source, relationship.id]));
    const relationships = [];
    for (const bulletImage of owner.references || []) {
      const source = pictureBulletSource(bulletImage);
      if (!source) throw new Error("Presentation picture bullet requires an embedded data URL or external URI before export.");
      if (relationshipIds.has(source)) continue;
      let target;
      let targetMode;
      if (bulletImage.uri) {
        target = source;
        targetMode = "External";
      } else {
        let media = mediaBySource.get(source);
        if (!media) {
          const data = decodeDataUrl?.(source);
          if (!data?.bytes || !data.extension) throw new TypeError("Presentation picture bullet has an unsupported embedded image data URL.");
          media = { imagePartId: nextMediaPartId++, source, ...data };
          mediaBySource.set(source, media);
          mediaParts.push(media);
        }
        target = `../media/image${media.imagePartId}.${media.extension}`;
      }
      const id = `rId${nextRelationshipIndex++}`;
      relationshipIds.set(source, id);
      relationships.push({ id, type: PRESENTATION_IMAGE_RELATIONSHIP_TYPE, target, ...(targetMode ? { targetMode } : {}) });
    }
    byOwner.set(owner.key, { relationshipIds, relationships, nextRelationshipIndex });
  }
  return { byOwner, mediaParts };
}

function imageContentTypeFromPath(target) {
  const extension = path.posix.extname(target).slice(1).toLowerCase();
  if (extension === "jpg" || extension === "jpeg") return "image/jpeg";
  if (extension === "png") return "image/png";
  if (extension === "gif") return "image/gif";
  if (extension === "svg") return "image/svg+xml";
  throw new TypeError(`Presentation picture bullet uses unsupported image extension ${extension || "(none)"}.`);
}

async function resolvePictureBulletImage(bulletImage, context) {
  if (!bulletImage?.relationshipId) return bulletImage;
  const relationship = (context.relationships || []).find((item) => item.id === bulletImage.relationshipId);
  if (!relationship?.type?.endsWith("/image")) throw new Error(`Presentation picture bullet references missing image relationship ${bulletImage.relationshipId}.`);
  if (relationship.targetMode?.toLowerCase() === "external" || bulletImage.relationshipMode === "link") {
    return { uri: relationship.target, relationshipMode: "link", ...(bulletImage.alt == null ? {} : { alt: bulletImage.alt }) };
  }
  const target = context.resolveTarget?.(context.partPath, relationship.target);
  if (!target) throw new Error(`Presentation picture bullet relationship ${bulletImage.relationshipId} has an invalid target.`);
  const bytes = await context.readPart?.(target);
  if (!bytes) throw new Error(`Presentation picture bullet relationship ${bulletImage.relationshipId} targets missing part ${target}.`);
  return {
    dataUrl: `data:${imageContentTypeFromPath(target)};base64,${Buffer.from(bytes).toString("base64")}`,
    relationshipMode: "embed",
    ...(bulletImage.alt == null ? {} : { alt: bulletImage.alt }),
  };
}

export async function resolvePresentationPictureBulletParagraphs(paragraphs = [], context = {}) {
  return Promise.all(paragraphs.map(async (paragraph) => paragraph?.bulletImage
    ? { ...paragraph, bulletImage: await resolvePictureBulletImage(paragraph.bulletImage, context) }
    : paragraph));
}

export async function resolvePresentationPictureBulletStyles(styles = {}, context = {}) {
  return Object.fromEntries(await Promise.all(Object.entries(styles || {}).map(async ([level, style]) => [
    level,
    style?.bulletImage ? { ...style, bulletImage: await resolvePictureBulletImage(style.bulletImage, context) } : style,
  ])));
}

export async function resolvePresentationPictureBulletMasterStyles(stylesByKind = {}, context = {}) {
  return Object.fromEntries(await Promise.all(Object.entries(stylesByKind || {}).map(async ([kind, styles]) => [
    kind,
    await resolvePresentationPictureBulletStyles(styles, context),
  ])));
}
