export function noLine() {
  return { fill: "transparent", width: 0 };
}

export function addDocumentStyles(document) {
  for (const [id, style] of Object.entries({
    TemplateTitle: {
      name: "Template Title",
      type: "paragraph",
      basedOn: "Normal",
      fontFamily: "Aptos Display",
      fontSize: 24,
      bold: true,
      color: "#0F3D4C",
      spaceAfterTwips: 90,
      keepNext: true,
    },
    TemplateSubtitle: {
      name: "Template Subtitle",
      type: "paragraph",
      basedOn: "Normal",
      fontFamily: "Aptos",
      fontSize: 10.5,
      color: "#4B6470",
      spaceAfterTwips: 150,
      keepNext: true,
    },
    TemplateHeading: {
      name: "Template Heading",
      type: "paragraph",
      basedOn: "Normal",
      fontFamily: "Aptos Display",
      fontSize: 13,
      bold: true,
      color: "#0F3D4C",
      spaceBeforeTwips: 160,
      spaceAfterTwips: 65,
      keepNext: true,
    },
    TemplateCallout: {
      name: "Template Callout",
      type: "paragraph",
      basedOn: "Normal",
      fontFamily: "Aptos",
      fontSize: 10,
      bold: true,
      color: "#0F5F61",
      spaceBeforeTwips: 70,
      spaceAfterTwips: 110,
      keepNext: true,
    },
    TableGrid: {
      name: "Table Grid",
      type: "table",
      fontFamily: "Aptos",
      fontSize: 9,
    },
  })) document.styles.add(id, style);
}

export function addPresentationText(slide, config) {
  return slide.shapes.add({
    geometry: "rect",
    fill: "transparent",
    line: noLine(),
    ...config,
  });
}

function mergedTitleValues(text, columns) {
  return [[text, ...Array(Math.max(0, columns - 1)).fill(null)]];
}

export function titleRange(sheet, range, text, columns = 5) {
  sheet.getRange(range).values = mergedTitleValues(text, columns);
  sheet.getRange(range).merge();
  sheet.getRange(range).format = {
    fill: "#0F3D4C",
    font: { bold: true, color: "#FFFFFF" },
    alignment: { horizontal: "left", vertical: "center" },
    rowHeightPx: 30,
  };
}

export function sectionRange(sheet, range, text, columns = 5) {
  sheet.getRange(range).values = mergedTitleValues(text, columns);
  sheet.getRange(range).merge();
  sheet.getRange(range).format = {
    fill: "#D9F0EE",
    font: { bold: true, color: "#0F3D4C" },
    alignment: { horizontal: "left", vertical: "center" },
    rowHeightPx: 22,
  };
}
