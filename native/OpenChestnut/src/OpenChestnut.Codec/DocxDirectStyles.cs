using DocumentFormat.OpenXml.Packaging;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Source-free authoring cannot rely on an application-provided style catalog.
// Define only the built-in table style that the public DocumentModel uses by
// default; arbitrary style graphs remain outside the bounded wire slice.
internal static class DocxDirectStyles
{
    internal static void AddRequiredStyles(MainDocumentPart mainPart, DocumentArtifact document)
    {
        if (!document.Blocks.Any(block =>
                block.ContentCase == DocumentBlock.ContentOneofCase.Table &&
                block.StyleId.Equals("TableGrid", StringComparison.Ordinal))) return;

        var borders = new W.TableBorders(
            Border(new W.TopBorder()),
            Border(new W.LeftBorder()),
            Border(new W.BottomBorder()),
            Border(new W.RightBorder()),
            Border(new W.InsideHorizontalBorder()),
            Border(new W.InsideVerticalBorder()));
        var style = new W.Style(
            new W.StyleName { Val = "Table Grid" },
            new W.StyleTableProperties(borders))
        {
            Type = W.StyleValues.Table,
            StyleId = "TableGrid",
        };
        var part = mainPart.AddNewPart<StyleDefinitionsPart>();
        part.Styles = new W.Styles(style);
        part.Styles.Save();
    }

    private static T Border<T>(T border) where T : W.BorderType
    {
        border.Val = W.BorderValues.Single;
        border.Color = "auto";
        border.Size = 4;
        border.Space = 0;
        return border;
    }
}
