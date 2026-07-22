using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

internal static class DocxSectionCodec
{
    internal static bool TryReadBoundary(W.Paragraph paragraph, out DocumentSection section, out bool editable)
    {
        section = new DocumentSection();
        editable = false;
        var properties = paragraph.ParagraphProperties;
        var native = properties?.SectionProperties;
        if (native is null) return false;
        if (paragraph.ChildElements.Any(child => child is not W.ParagraphProperties)) return false;
        section = Read(native);
        editable = properties!.ChildElements.All(child => child is W.ParagraphStyleId or W.SectionProperties) &&
                   IsBounded(native);
        return true;
    }

    internal static DocumentSection Read(W.SectionProperties source)
    {
        var size = source.GetFirstChild<W.PageSize>();
        var margins = source.GetFirstChild<W.PageMargin>();
        return new DocumentSection
        {
            BreakType = FromNativeBreak(source.GetFirstChild<W.SectionType>()?.Val?.Value),
            PageWidthTwips = size?.Width?.Value ?? 12240U,
            PageHeightTwips = size?.Height?.Value ?? 15840U,
            Landscape = size?.Orient?.Value == W.PageOrientationValues.Landscape,
            MarginTopTwips = Positive(margins?.Top?.Value, 1440U),
            MarginRightTwips = margins?.Right?.Value ?? 1440U,
            MarginBottomTwips = Positive(margins?.Bottom?.Value, 1440U),
            MarginLeftTwips = margins?.Left?.Value ?? 1440U,
            MarginGutterTwips = margins?.Gutter?.Value ?? 0U,
        };
    }

    internal static W.Paragraph BuildBoundary(
        DocumentSection source,
        IEnumerable<OpenXmlElement> references,
        bool differentFirstPage,
        bool gutterAtTop)
    {
        Validate(source, "Document section", gutterAtTop);
        return new W.Paragraph(new W.ParagraphProperties(BuildProperties(source, references, differentFirstPage)));
    }

    internal static W.SectionProperties BuildFinal(
        IEnumerable<OpenXmlElement> references,
        bool differentFirstPage) =>
        BuildProperties(Default(), references, differentFirstPage);

    internal static void Apply(W.Paragraph paragraph, DocumentSection requested, bool gutterAtTop)
    {
        Validate(requested, "Document section", gutterAtTop);
        var properties = paragraph.ParagraphProperties ?? throw new CodecException(
            "document_source_binding_mismatch", "Source section boundary has no paragraph properties.", "word/document.xml");
        var native = properties.SectionProperties ?? throw new CodecException(
            "document_source_binding_mismatch", "Source section boundary has no section properties.", "word/document.xml");
        if (!IsBounded(native))
            throw new CodecException("unsupported_document_edit", "Section contains unmodeled properties and cannot be edited safely.", "word/document.xml");

        Replace(native, native.GetFirstChild<W.SectionType>(), BuildType(requested.BreakType));
        Replace(native, native.GetFirstChild<W.PageSize>(), BuildPageSize(requested));
        var sourceMargins = native.GetFirstChild<W.PageMargin>();
        Replace(native, sourceMargins, BuildPageMargin(requested, sourceMargins));
    }

    internal static string ResidualHash(W.Paragraph paragraph)
    {
        var clone = (W.Paragraph)paragraph.CloneNode(true);
        var section = clone.ParagraphProperties?.SectionProperties;
        section?.GetFirstChild<W.SectionType>()?.Remove();
        section?.GetFirstChild<W.PageSize>()?.Remove();
        section?.GetFirstChild<W.PageMargin>()?.Remove();
        return Hash(clone.OuterXml);
    }

    internal static void Validate(DocumentSection section, string label, bool gutterAtTop)
    {
        if (section.BreakType == DocumentSectionBreak.Unspecified)
            throw new CodecException("invalid_document_section", $"{label} requires a supported break type.");
        if (section.PageWidthTwips is < 1 or > 31680 || section.PageHeightTwips is < 1 or > 31680)
            throw new CodecException("invalid_document_section", $"{label} page size must be 1 through 31680 twentieths of a point.");
        foreach (var (name, value) in new[]
                 {
                     ("top", section.MarginTopTwips), ("right", section.MarginRightTwips),
                     ("bottom", section.MarginBottomTwips), ("left", section.MarginLeftTwips),
                     ("gutter", section.MarginGutterTwips),
                 })
            if (value > 31680)
                throw new CodecException("invalid_document_section", $"{label} {name} margin exceeds 31680 twentieths of a point.");
        var horizontalGutter = gutterAtTop ? 0U : section.MarginGutterTwips;
        var verticalGutter = gutterAtTop ? section.MarginGutterTwips : 0U;
        if ((ulong)section.MarginLeftTwips + section.MarginRightTwips + horizontalGutter >= section.PageWidthTwips ||
            (ulong)section.MarginTopTwips + section.MarginBottomTwips + verticalGutter >= section.PageHeightTwips)
            throw new CodecException("invalid_document_section", $"{label} margins and binding gutter must leave a positive page content area.");
    }

    private static W.SectionProperties BuildProperties(
        DocumentSection source,
        IEnumerable<OpenXmlElement> references,
        bool differentFirstPage)
    {
        var properties = new W.SectionProperties();
        foreach (var reference in references) properties.Append(reference.CloneNode(true));
        properties.Append(BuildType(source.BreakType), BuildPageSize(source), BuildPageMargin(source));
        if (differentFirstPage) properties.Append(new W.TitlePage());
        return properties;
    }

    private static W.SectionType BuildType(DocumentSectionBreak value) => new()
    {
        Val = value switch
        {
            DocumentSectionBreak.Continuous => W.SectionMarkValues.Continuous,
            DocumentSectionBreak.EvenPage => W.SectionMarkValues.EvenPage,
            DocumentSectionBreak.OddPage => W.SectionMarkValues.OddPage,
            _ => W.SectionMarkValues.NextPage,
        },
    };

    private static W.PageSize BuildPageSize(DocumentSection source) => new()
    {
        Width = source.PageWidthTwips,
        Height = source.PageHeightTwips,
        Orient = source.Landscape ? W.PageOrientationValues.Landscape : W.PageOrientationValues.Portrait,
    };

    private static W.PageMargin BuildPageMargin(DocumentSection source, W.PageMargin? sourceMargins = null) => new()
    {
        Top = checked((int)source.MarginTopTwips),
        Right = source.MarginRightTwips,
        Bottom = checked((int)source.MarginBottomTwips),
        Left = source.MarginLeftTwips,
        Header = sourceMargins?.Header?.Value ?? 720U,
        Footer = sourceMargins?.Footer?.Value ?? 720U,
        Gutter = source.MarginGutterTwips,
    };

    private static DocumentSection Default() => new()
    {
        BreakType = DocumentSectionBreak.NextPage,
        PageWidthTwips = 12240,
        PageHeightTwips = 15840,
        MarginTopTwips = 1440,
        MarginRightTwips = 1440,
        MarginBottomTwips = 1440,
        MarginLeftTwips = 1440,
        MarginGutterTwips = 0,
    };

    private static bool IsBounded(W.SectionProperties source) => source.ChildElements.All(child =>
        child is W.HeaderReference or W.FooterReference or W.SectionType or W.PageSize or W.PageMargin or W.TitlePage);

    private static DocumentSectionBreak FromNativeBreak(W.SectionMarkValues? value) =>
        value == W.SectionMarkValues.Continuous ? DocumentSectionBreak.Continuous :
        value == W.SectionMarkValues.EvenPage ? DocumentSectionBreak.EvenPage :
        value == W.SectionMarkValues.OddPage ? DocumentSectionBreak.OddPage :
        DocumentSectionBreak.NextPage;

    private static uint Positive(int? value, uint fallback) => value is >= 0 ? checked((uint)value.Value) : fallback;

    private static void Replace(W.SectionProperties owner, OpenXmlElement? current, OpenXmlElement replacement)
    {
        if (current is null)
        {
            var anchor = owner.ChildElements.FirstOrDefault(child => child is W.TitlePage);
            if (anchor is null) owner.Append(replacement);
            else owner.InsertBefore(replacement, anchor);
        }
        else current.InsertAfterSelf(replacement);
        current?.Remove();
    }

    private static string Hash(string value) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();
}
