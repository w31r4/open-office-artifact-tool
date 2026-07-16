using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

internal static class DocxDirectStyles
{
    private static readonly DocumentStyle[] BuiltIns =
    [
        Style("Normal", "Normal", DocumentStyleType.Paragraph, 22, "Aptos"),
        Style("Title", "Title", DocumentStyleType.Paragraph, 48, "Aptos Display", bold: true),
        Style("Heading1", "Heading 1", DocumentStyleType.Paragraph, 32, "Aptos Display", bold: true),
        Style("Heading2", "Heading 2", DocumentStyleType.Paragraph, 26, "Aptos", bold: true),
    ];

    internal static void Read(MainDocumentPart mainPart, DocumentArtifact document)
    {
        var styles = mainPart.StyleDefinitionsPart?.Styles;
        if (styles is not null)
        {
            document.DefaultRunStyle = DocxFormattingCodec.ReadBaseRunFormatting(
                styles.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle);
            foreach (var source in styles.Elements<W.Style>())
            {
                var type = FromNativeType(source.Type?.Value);
                if (type == DocumentStyleType.Unspecified || string.IsNullOrWhiteSpace(source.StyleId?.Value)) continue;
                var style = new DocumentStyle
                {
                    Id = source.StyleId!.Value!,
                    Name = source.StyleName?.Val?.Value ?? source.StyleId.Value!,
                    Type = type,
                    BasedOn = source.BasedOn?.Val?.Value ?? string.Empty,
                    RunFormat = DocxFormattingCodec.ReadStyleRunFormatting(source.StyleRunProperties),
                    ParagraphFormat = DocxFormattingCodec.ReadStyleParagraphFormatting(source.StyleParagraphProperties),
                };
                document.Styles.Add(style);
            }
        }
        var ids = document.Styles.Select(style => style.Id).ToHashSet(StringComparer.Ordinal);
        foreach (var builtIn in BuiltIns)
            if (ids.Add(builtIn.Id)) document.Styles.Add(builtIn.Clone());
    }

    internal static void AddRequiredStyles(MainDocumentPart mainPart, DocumentArtifact document)
    {
        Validate(document);
        var requested = document.Styles.Select(style => style.Clone()).ToList();
        var ids = requested.Select(style => style.Id).ToHashSet(StringComparer.Ordinal);
        foreach (var builtIn in BuiltIns)
            if (!ids.Contains(builtIn.Id)) requested.Add(builtIn.Clone());

        var needsTableGrid = document.Blocks.Any(block =>
            block.ContentCase == DocumentBlock.ContentOneofCase.Table &&
            block.StyleId.Equals("TableGrid", StringComparison.Ordinal));
        if (needsTableGrid && ids.Add("TableGrid"))
            requested.Add(new DocumentStyle
            {
                Id = "TableGrid",
                Name = "Table Grid",
                Type = DocumentStyleType.Table,
            });

        var root = new W.Styles();
        var defaults = DocxFormattingCodec.BuildDefaultRunProperties(document.DefaultRunStyle);
        if (defaults is not null)
            root.Append(new W.DocDefaults(new W.RunPropertiesDefault(defaults)));
        foreach (var source in requested) root.Append(BuildStyle(source));
        var part = mainPart.AddNewPart<StyleDefinitionsPart>();
        part.Styles = root;
        root.Save();
    }

    internal static void AssertSourceUnchanged(MainDocumentPart mainPart, DocumentArtifact requested)
    {
        var source = new DocumentArtifact();
        Read(mainPart, source);
        var sourceStyles = source.Styles.OrderBy(style => style.Id, StringComparer.Ordinal).ToArray();
        var requestedStyles = requested.Styles.OrderBy(style => style.Id, StringComparer.Ordinal).ToArray();
        if (!Equals(source.DefaultRunStyle, requested.DefaultRunStyle) ||
            sourceStyles.Length != requestedStyles.Length ||
            sourceStyles.Where((style, index) => !style.Equals(requestedStyles[index])).Any())
            throw new CodecException(
                "unsupported_document_style_edit",
                "Source-preserving DOCX export does not permit changing the imported style catalog or document defaults.",
                "word/styles.xml");
    }

    internal static void Validate(DocumentArtifact document, bool allowImportedCycles = false)
    {
        DocxFormattingCodec.Validate(document.DefaultRunStyle, "Document default run style");
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var style in document.Styles)
        {
            if (string.IsNullOrWhiteSpace(style.Id) || style.Id.Length > 253 || style.Id.Any(char.IsControl))
                throw new CodecException("invalid_document_style", "Document style IDs must contain 1 through 253 characters without controls.");
            if (!ids.Add(style.Id))
                throw new CodecException("invalid_document_style", $"Document contains duplicate style ID {style.Id}.");
            if (style.Type == DocumentStyleType.Unspecified)
                throw new CodecException("invalid_document_style", $"Document style {style.Id} requires paragraph, character, or table type.");
            if (string.IsNullOrWhiteSpace(style.Name) || style.Name.Length > 255 || style.Name.Any(char.IsControl))
                throw new CodecException("invalid_document_style", $"Document style {style.Id} name must contain 1 through 255 characters without controls.");
            if (style.BasedOn.Length > 253 || style.BasedOn.Any(char.IsControl))
                throw new CodecException("invalid_document_style", $"Document style {style.Id} based-on ID is invalid.");
            DocxFormattingCodec.Validate(style.RunFormat, $"Document style {style.Id}");
            DocxFormattingCodec.Validate(style.ParagraphFormat, $"Document style {style.Id}");
            if (style.Type != DocumentStyleType.Paragraph && DocxFormattingCodec.HasParagraphFormatting(style.ParagraphFormat))
                throw new CodecException("invalid_document_style", $"Document style {style.Id} paragraph formatting requires paragraph type.");
        }
        foreach (var style in document.Styles)
            if (!string.IsNullOrWhiteSpace(style.BasedOn) && !ids.Contains(style.BasedOn) &&
                BuiltIns.All(item => !item.Id.Equals(style.BasedOn, StringComparison.Ordinal)))
                throw new CodecException("invalid_document_style", $"Document style {style.Id} is based on missing style {style.BasedOn}.");
        if (!allowImportedCycles)
            foreach (var style in document.Styles)
            {
                var seen = new HashSet<string>(StringComparer.Ordinal) { style.Id };
                var parent = style.BasedOn;
                while (!string.IsNullOrWhiteSpace(parent) && ids.Contains(parent))
                {
                    if (!seen.Add(parent))
                        throw new CodecException("invalid_document_style", $"Document style {style.Id} has a cyclic based-on chain.");
                    parent = document.Styles.First(item => item.Id == parent).BasedOn;
                }
            }
    }

    private static W.Style BuildStyle(DocumentStyle source)
    {
        var style = new W.Style
        {
            Type = source.Type switch
            {
                DocumentStyleType.Character => W.StyleValues.Character,
                DocumentStyleType.Table => W.StyleValues.Table,
                _ => W.StyleValues.Paragraph,
            },
            StyleId = source.Id,
            Default = source.Id.Equals("Normal", StringComparison.Ordinal) ? true : null,
        };
        style.Append(new W.StyleName { Val = source.Name });
        if (!string.IsNullOrWhiteSpace(source.BasedOn)) style.Append(new W.BasedOn { Val = source.BasedOn });
        var paragraph = DocxFormattingCodec.BuildStyleParagraphProperties(source.ParagraphFormat, $"Document style {source.Id}");
        if (paragraph is not null) style.Append(paragraph);
        var run = DocxFormattingCodec.BuildStyleRunProperties(source.RunFormat, $"Document style {source.Id}");
        if (run is not null) style.Append(run);
        if (source.Id.Equals("TableGrid", StringComparison.Ordinal) && source.Type == DocumentStyleType.Table)
            style.Append(new W.StyleTableProperties(new W.TableBorders(
                Border(new W.TopBorder()), Border(new W.LeftBorder()), Border(new W.BottomBorder()),
                Border(new W.RightBorder()), Border(new W.InsideHorizontalBorder()), Border(new W.InsideVerticalBorder()))));
        return style;
    }

    private static DocumentStyle Style(string id, string name, DocumentStyleType type, uint halfPoints, string font, bool bold = false)
    {
        var formatting = new DocumentRunFormatting { FontFamily = font, FontSizeHalfPoints = halfPoints };
        if (bold) formatting.Bold = true;
        return new DocumentStyle { Id = id, Name = name, Type = type, RunFormat = formatting };
    }

    private static DocumentStyleType FromNativeType(W.StyleValues? value) =>
        value == W.StyleValues.Paragraph ? DocumentStyleType.Paragraph :
        value == W.StyleValues.Character ? DocumentStyleType.Character :
        value == W.StyleValues.Table ? DocumentStyleType.Table :
        DocumentStyleType.Unspecified;

    private static T Border<T>(T border) where T : W.BorderType
    {
        border.Val = W.BorderValues.Single;
        border.Color = "auto";
        border.Size = 4;
        border.Space = 0;
        return border;
    }
}
