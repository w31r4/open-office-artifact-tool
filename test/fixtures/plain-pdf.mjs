function escapePdfLiteral(value) {
  const text = String(value);
  if (/[^\x20-\x7e]/u.test(text)) throw new Error("plain PDF fixture text must be printable ASCII");
  return text.replace(/[\\()]/gu, "\\$&");
}

function pdfFromIndirectObjects(objects) {
  let body = "%PDF-1.7\n";
  const offsets = [0];
  for (let index = 1; index < objects.length; index += 1) {
    offsets[index] = Buffer.byteLength(body);
    body += `${index} 0 obj\n${objects[index]}\nendobj\n`;
  }
  const xref = Buffer.byteLength(body);
  body += `xref\n0 ${objects.length}\n0000000000 65535 f \n`;
  for (let index = 1; index < objects.length; index += 1) {
    body += `${String(offsets[index]).padStart(10, "0")} 00000 n \n`;
  }
  body += `trailer\n<< /Size ${objects.length} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return new Uint8Array(Buffer.from(body, "latin1"));
}

export function plainPdfBytes(pages = ["First page", "Second page"]) {
  if (!Array.isArray(pages) || !pages.length) throw new Error("plain PDF fixture requires at least one page");
  const descriptors = pages.map((page) => typeof page === "string" ? { text: page } : page);
  const objects = [
    undefined,
    "<< /Type /Catalog /Pages 2 0 R >>",
    "",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
  ];
  const pageRefs = [];
  for (const [index, descriptor] of descriptors.entries()) {
    const width = Number(descriptor.width ?? 612);
    const height = Number(descriptor.height ?? 792);
    const rotation = Number(descriptor.rotation ?? 0);
    if (!Number.isFinite(width) || width <= 96 || !Number.isFinite(height) || height <= 132) throw new Error("plain PDF fixture page size must fit its visible test content");
    if (![0, 90, 180, 270].includes(rotation)) throw new Error("plain PDF fixture rotation must be a right angle");
    const pageObject = objects.length;
    const contentObject = pageObject + 1;
    pageRefs.push(`${pageObject} 0 R`);
    const shade = (0.94 - (index % 3) * 0.04).toFixed(2);
    const content = [
      `q ${shade} ${shade} 1 rg 48 48 ${width - 96} ${height - 96} re f Q`,
      `BT /F1 24 Tf 72 ${height - 96} Td (${escapePdfLiteral(descriptor.text ?? `Page ${index + 1}`)}) Tj ET`,
      `BT /F1 12 Tf 72 ${height - 132} Td (fixture-page-${index + 1}) Tj ET`,
    ].join("\n");
    objects.push(
      `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${width} ${height}] /Rotate ${rotation} /Resources << /Font << /F1 3 0 R >> >> /Contents ${contentObject} 0 R >>`,
      `<< /Length ${Buffer.byteLength(content, "latin1")} >>\nstream\n${content}\nendstream`,
    );
  }
  objects[2] = `<< /Type /Pages /Kids [${pageRefs.join(" ")}] /Count ${pageRefs.length} >>`;
  return pdfFromIndirectObjects(objects);
}
