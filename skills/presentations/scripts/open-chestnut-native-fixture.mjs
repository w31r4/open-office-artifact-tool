import JSZip from "jszip";

// Builds a standards-valid, source-bound native-object graph for the runnable
// OpenChestnut fixture. The graph deliberately crosses an embedded package,
// a SmartArt four-part root, and a recursive contentPart/customXmlProps edge.
function removeFirstPlaceholderTransform(xml, partPath, index) {
  let removed = false;
  const output = String(xml || "").replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, (shape) => {
    const placeholder = /<p:ph\b[^>]*\/?\s*>/.exec(shape)?.[0];
    if (removed || !placeholder) return shape;
    if (index != null && Number(/\bidx="(\d+)"/.exec(placeholder)?.[1] || 0) !== Number(index)) return shape;
    const next = shape.replace(/<a:xfrm\b[^>]*>[\s\S]*?<\/a:xfrm>/, "");
    removed = next !== shape;
    return next;
  });
  if (!removed) throw new Error(`OpenChestnut native fixture could not remove a placeholder transform from ${partPath}.`);
  return output;
}

export async function addOpenChestnutNativeGraphFixture(bytes, embeddedWorkbookBytes, options = {}) {
  const zip = await JSZip.loadAsync(bytes);
  const namespaces = ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const ole = `<p:graphicFrame${namespaces}><p:nvGraphicFramePr><p:cNvPr id="100" name="Embedded workbook"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="11191875" y="2190750"/><a:ext cx="762000" cy="762000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj showAsIcon="1" r:id="rIdNativeOle" imgW="762000" imgH="762000" progId="Excel.Sheet.12"><p:embed/><p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdNativePreview"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="11191875" y="2190750"/><a:ext cx="762000" cy="762000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>`;
  const diagram = `<p:graphicFrame${namespaces}><p:nvGraphicFramePr><p:cNvPr id="101" name="Preserved diagram"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="11191875" y="3143250"/><a:ext cx="762000" cy="762000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rIdNativeDm" r:lo="rIdNativeLo" r:qs="rIdNativeQs" r:cs="rIdNativeCs"/></a:graphicData></a:graphic></p:graphicFrame>`;
  const content = `<p:grpSp${namespaces}><p:nvGrpSpPr><p:cNvPr id="102" name="Native content group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="11191875" y="4095750"/><a:ext cx="762000" cy="762000"/><a:chOff x="0" y="0"/><a:chExt cx="762000" cy="762000"/></a:xfrm></p:grpSpPr><p:contentPart r:id="rIdNativeContent"/></p:grpSp>`;
  const slidePath = "ppt/slides/slide1.xml";
  const slideXml = await zip.file(slidePath)?.async("text");
  if (!slideXml?.includes("</p:spTree>")) throw new Error(`OpenChestnut native fixture requires ${slidePath} with a p:spTree.`);
  zip.file(slidePath, slideXml.replace("</p:spTree>", `${ole}${diagram}${content}</p:spTree>`));

  if (options.removeMasterPlaceholderFrame) {
    const masterPath = "ppt/slideMasters/slideMaster1.xml";
    zip.file(masterPath, removeFirstPlaceholderTransform(await zip.file(masterPath)?.async("text"), masterPath));
  }
  if (options.removeSlidePlaceholderFrameIndex != null) {
    zip.file(slidePath, removeFirstPlaceholderTransform(await zip.file(slidePath)?.async("text"), slidePath, options.removeSlidePlaceholderFrameIndex));
  }

  const relationshipsPath = "ppt/slides/_rels/slide1.xml.rels";
  const relationshipsXml = await zip.file(relationshipsPath)?.async("text");
  if (!relationshipsXml?.includes("</Relationships>")) throw new Error(`OpenChestnut native fixture requires ${relationshipsPath}.`);
  const nativeRelationships = '<Relationship Id="rIdNativeOle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/native-workbook.xlsx"/><Relationship Id="rIdNativePreview" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/native-preview.png"/><Relationship Id="rIdNativeDm" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/native-data.xml"/><Relationship Id="rIdNativeLo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="../diagrams/native-layout.xml"/><Relationship Id="rIdNativeQs" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="../diagrams/native-style.xml"/><Relationship Id="rIdNativeCs" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="../diagrams/native-colors.xml"/><Relationship Id="rIdNativeContent" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/native-content.xml"/>';
  zip.file(relationshipsPath, relationshipsXml.replace("</Relationships>", `${nativeRelationships}</Relationships>`));

  const contentTypesPath = "[Content_Types].xml";
  const contentTypes = await zip.file(contentTypesPath)?.async("text");
  if (!contentTypes?.includes("</Types>")) throw new Error("OpenChestnut native fixture requires [Content_Types].xml.");
  const overrides = '<Override PartName="/ppt/embeddings/native-workbook.xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/><Override PartName="/ppt/media/native-preview.png" ContentType="image/png"/><Override PartName="/ppt/diagrams/native-data.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/><Override PartName="/ppt/diagrams/native-layout.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"/><Override PartName="/ppt/diagrams/native-style.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml"/><Override PartName="/ppt/diagrams/native-colors.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"/><Override PartName="/ppt/customXml/native-content.xml" ContentType="application/xml"/><Override PartName="/ppt/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>';
  zip.file(contentTypesPath, contentTypes.replace("</Types>", `${overrides}</Types>`));
  if (!embeddedWorkbookBytes?.length) throw new Error("OpenChestnut native fixture requires a valid embedded XLSX workbook.");
  zip.file("ppt/embeddings/native-workbook.xlsx", embeddedWorkbookBytes);
  zip.file("ppt/media/native-preview.png", "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=", { base64: true });
  zip.file("ppt/diagrams/native-data.xml", '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:ptLst/><dgm:cxnLst/><dgm:bg/><dgm:whole/></dgm:dataModel>');
  zip.file("ppt/diagrams/native-layout.xml", '<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:open-office:native-layout"><dgm:title val="Native"/><dgm:desc val="Native layout"/><dgm:catLst/><dgm:layoutNode name="root"/></dgm:layoutDef>');
  zip.file("ppt/diagrams/native-style.xml", '<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:open-office:native-style"><dgm:title val="Native"/><dgm:desc val="Native style"/><dgm:catLst/><dgm:styleLbl name="node0"/></dgm:styleDef>');
  zip.file("ppt/diagrams/native-colors.xml", '<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:open-office:native-colors"><dgm:title val="Native"/><dgm:desc val="Native colors"/><dgm:catLst/></dgm:colorsDef>');
  zip.file("ppt/customXml/native-content.xml", '<native xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:link="rIdPayload">preserve me</native>');
  zip.file("ppt/customXml/_rels/native-content.xml.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPayload" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>');
  zip.file("ppt/customXml/itemProps1.xml", '<ds:datastoreItem ds:itemID="{00112233-4455-6677-8899-AABBCCDDEEFF}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>');
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}
