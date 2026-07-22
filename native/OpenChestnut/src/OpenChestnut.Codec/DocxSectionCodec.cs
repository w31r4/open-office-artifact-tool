using System.Globalization;
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
        var result = new DocumentSection
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
        if (TryReadColumns(source, out var columns) && columns is not null) result.Columns = columns;
        if (TryReadPageNumbering(source, out var pageNumbering) && pageNumbering is not null) result.PageNumbering = pageNumbering;
        if (TryReadLineNumbering(source, out var lineNumbering) && lineNumbering is not null) result.LineNumbering = lineNumbering;
        return result;
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
        var sourceColumns = native.GetFirstChild<W.Columns>();
        if (requested.Columns is null) sourceColumns?.Remove();
        else Replace(native, sourceColumns, BuildColumns(requested.Columns));
        var sourcePageNumbering = native.GetFirstChild<W.PageNumberType>();
        if (requested.PageNumbering is null) sourcePageNumbering?.Remove();
        else Replace(native, sourcePageNumbering, BuildPageNumbering(requested.PageNumbering));
        var sourceLineNumbering = native.GetFirstChild<W.LineNumberType>();
        if (requested.LineNumbering is null) sourceLineNumbering?.Remove();
        else Replace(native, sourceLineNumbering, BuildLineNumbering(requested.LineNumbering));
    }

    internal static string ResidualHash(W.Paragraph paragraph)
    {
        var clone = (W.Paragraph)paragraph.CloneNode(true);
        var section = clone.ParagraphProperties?.SectionProperties;
        section?.GetFirstChild<W.SectionType>()?.Remove();
        section?.GetFirstChild<W.PageSize>()?.Remove();
        section?.GetFirstChild<W.PageMargin>()?.Remove();
        section?.GetFirstChild<W.LineNumberType>()?.Remove();
        section?.GetFirstChild<W.PageNumberType>()?.Remove();
        section?.GetFirstChild<W.Columns>()?.Remove();
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
        if (section.Columns is not null)
        {
            var availableWidth = (ulong)section.PageWidthTwips - section.MarginLeftTwips - section.MarginRightTwips - horizontalGutter;
            if (section.Columns.Definitions.Count > 0)
            {
                if (section.Columns.Count != 0 || section.Columns.SpacingTwips != 0)
                    throw new CodecException("invalid_document_section", $"{label} custom-width columns cannot combine definitions with equal-width count or spacing.");
                if (section.Columns.Definitions.Count > 45)
                    throw new CodecException("invalid_document_section", $"{label} custom-width columns require 1 through 45 definitions.");
                ulong occupiedWidth = 0;
                foreach (var definition in section.Columns.Definitions)
                {
                    if (definition.WidthTwips is < 1 or > 31680)
                        throw new CodecException("invalid_document_section", $"{label} custom column widths must be 1 through 31680 twentieths of a point.");
                    if (definition.SpacingAfterTwips > 31680)
                        throw new CodecException("invalid_document_section", $"{label} custom column spacing must not exceed 31680 twentieths of a point.");
                    occupiedWidth += (ulong)definition.WidthTwips + definition.SpacingAfterTwips;
                }
                if (occupiedWidth > availableWidth)
                    throw new CodecException("invalid_document_section", $"{label} custom column widths and spacing must fit within the page content width.");
            }
            else
            {
                if (section.Columns.Count is < 1 or > 45)
                    throw new CodecException("invalid_document_section", $"{label} equal-width column count must be 1 through 45.");
                if (section.Columns.SpacingTwips > 31680)
                    throw new CodecException("invalid_document_section", $"{label} column spacing must not exceed 31680 twentieths of a point.");
                if ((ulong)(section.Columns.Count - 1) * section.Columns.SpacingTwips >= availableWidth)
                    throw new CodecException("invalid_document_section", $"{label} column spacing must leave positive width for every text column.");
            }
        }
        if (section.PageNumbering is not null)
        {
            if (!section.PageNumbering.HasStart && section.PageNumbering.Format == DocumentSectionPageNumberFormat.Unspecified)
                throw new CodecException("invalid_document_section", $"{label} page numbering requires a start value or supported format.");
            if (section.PageNumbering.HasStart && section.PageNumbering.Start > int.MaxValue)
                throw new CodecException("invalid_document_section", $"{label} page-number start must not exceed 2147483647.");
            if (!IsSupportedPageNumberFormat(section.PageNumbering.Format))
                throw new CodecException("invalid_document_section", $"{label} page-number format is unsupported.");
        }
        if (section.LineNumbering is not null)
        {
            if (section.LineNumbering.CountBy is < 1 or > 32767)
                throw new CodecException("invalid_document_section", $"{label} line-number countBy must be 1 through 32767.");
            if (section.LineNumbering.HasStart && section.LineNumbering.Start > 32767)
                throw new CodecException("invalid_document_section", $"{label} line-number start must be 0 through 32767.");
            if (section.LineNumbering.HasDistanceTwips && section.LineNumbering.DistanceTwips > 31680)
                throw new CodecException("invalid_document_section", $"{label} line-number distance must be 0 through 31680 twentieths of a point.");
            if (!IsSupportedLineNumberRestart(section.LineNumbering.Restart))
                throw new CodecException("invalid_document_section", $"{label} line-number restart is unsupported.");
        }
    }

    private static W.SectionProperties BuildProperties(
        DocumentSection source,
        IEnumerable<OpenXmlElement> references,
        bool differentFirstPage)
    {
        var properties = new W.SectionProperties();
        foreach (var reference in references) properties.Append(reference.CloneNode(true));
        properties.Append(BuildType(source.BreakType), BuildPageSize(source), BuildPageMargin(source));
        if (source.LineNumbering is not null) properties.Append(BuildLineNumbering(source.LineNumbering));
        if (source.PageNumbering is not null) properties.Append(BuildPageNumbering(source.PageNumbering));
        if (source.Columns is not null) properties.Append(BuildColumns(source.Columns));
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

    private static W.Columns BuildColumns(DocumentSectionColumns source)
    {
        var result = new W.Columns { Separator = source.Separator };
        if (source.Definitions.Count > 0)
        {
            result.EqualWidth = false;
            foreach (var definition in source.Definitions)
            {
                var column = new W.Column
                {
                    Width = definition.WidthTwips.ToString(CultureInfo.InvariantCulture),
                };
                if (definition.SpacingAfterTwips > 0)
                    column.Space = definition.SpacingAfterTwips.ToString(CultureInfo.InvariantCulture);
                result.Append(column);
            }
        }
        else
        {
            result.EqualWidth = true;
            result.ColumnCount = checked((short)source.Count);
            result.Space = source.SpacingTwips.ToString(CultureInfo.InvariantCulture);
        }
        return result;
    }

    private static W.PageNumberType BuildPageNumbering(DocumentSectionPageNumbering source)
    {
        var result = new W.PageNumberType();
        if (source.HasStart) result.Start = checked((int)source.Start);
        if (source.Format != DocumentSectionPageNumberFormat.Unspecified) result.Format = ToNativePageNumberFormat(source.Format);
        return result;
    }

    private static W.LineNumberType BuildLineNumbering(DocumentSectionLineNumbering source)
    {
        var result = new W.LineNumberType
        {
            CountBy = checked((short)source.CountBy),
        };
        if (source.HasStart) result.Start = checked((short)source.Start);
        if (source.HasDistanceTwips) result.Distance = source.DistanceTwips.ToString(CultureInfo.InvariantCulture);
        if (source.Restart != DocumentSectionLineNumberRestart.Unspecified) result.Restart = ToNativeLineNumberRestart(source.Restart);
        return result;
    }

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

    private static bool IsBounded(W.SectionProperties source)
    {
        if (!source.ChildElements.All(child => child is W.HeaderReference or W.FooterReference or
                W.SectionType or W.PageSize or W.PageMargin or W.LineNumberType or W.PageNumberType or W.Columns or W.TitlePage) ||
            source.Elements<W.SectionType>().Count() > 1 ||
            source.Elements<W.PageSize>().Count() > 1 ||
            source.Elements<W.PageMargin>().Count() > 1 ||
            source.Elements<W.LineNumberType>().Count() > 1 ||
            source.Elements<W.PageNumberType>().Count() > 1 ||
            source.Elements<W.TitlePage>().Count() > 1)
            return false;
        return TryReadColumns(source, out _) && TryReadPageNumbering(source, out _) && TryReadLineNumbering(source, out _);
    }

    private static bool TryReadColumns(W.SectionProperties source, out DocumentSectionColumns? result)
    {
        result = null;
        var matches = source.Elements<W.Columns>().ToArray();
        if (matches.Length == 0) return true;
        if (matches.Length != 1) return false;
        var columns = matches[0];
        if (columns.ExtendedAttributes.Any() || columns.NamespaceDeclarations.Any() || columns.MCAttributes is not null)
            return false;
        try
        {
            return columns.EqualWidth?.Value == false
                ? TryReadCustomWidthColumns(columns, out result)
                : TryReadEqualWidthColumns(columns, out result);
        }
        catch (Exception exception) when (exception is FormatException or InvalidOperationException or OverflowException)
        {
            return false;
        }
    }

    private static bool TryReadEqualWidthColumns(W.Columns columns, out DocumentSectionColumns? result)
    {
        result = null;
        if (columns.HasChildren) return false;
        var count = columns.ColumnCount?.Value ?? (short)1;
        if (count is < 1 or > 45 ||
            columns.Space?.Value is not { } spacingText ||
            !uint.TryParse(spacingText, NumberStyles.None, CultureInfo.InvariantCulture, out var spacing) ||
            spacing > 31680)
            return false;
        result = new DocumentSectionColumns
        {
            Count = checked((uint)count),
            SpacingTwips = spacing,
            Separator = columns.Separator?.Value ?? false,
        };
        return true;
    }

    private static bool TryReadCustomWidthColumns(W.Columns columns, out DocumentSectionColumns? result)
    {
        result = null;
        if (columns.ColumnCount is not null || columns.Space is not null) return false;
        var definitions = columns.Elements<W.Column>().ToArray();
        if (definitions.Length is < 1 or > 45 || columns.ChildElements.Count != definitions.Length) return false;
        var value = new DocumentSectionColumns { Separator = columns.Separator?.Value ?? false };
        foreach (var definition in definitions)
        {
            if (definition.HasChildren || definition.ExtendedAttributes.Any() || definition.NamespaceDeclarations.Any() || definition.MCAttributes is not null ||
                definition.Width?.Value is not { } widthText ||
                !uint.TryParse(widthText, NumberStyles.None, CultureInfo.InvariantCulture, out var width) ||
                width is < 1 or > 31680)
                return false;
            var spacing = 0U;
            if (definition.Space?.Value is { } spacingText &&
                (!uint.TryParse(spacingText, NumberStyles.None, CultureInfo.InvariantCulture, out spacing) || spacing > 31680))
                return false;
            value.Definitions.Add(new DocumentSectionColumnDefinition
            {
                WidthTwips = width,
                SpacingAfterTwips = spacing,
            });
        }
        result = value;
        return true;
    }

    private static bool TryReadPageNumbering(W.SectionProperties source, out DocumentSectionPageNumbering? result)
    {
        result = null;
        var matches = source.Elements<W.PageNumberType>().ToArray();
        if (matches.Length == 0) return true;
        if (matches.Length != 1) return false;
        var pageNumbering = matches[0];
        if (pageNumbering.HasChildren || pageNumbering.ExtendedAttributes.Any() || pageNumbering.NamespaceDeclarations.Any() ||
            pageNumbering.MCAttributes is not null || pageNumbering.ChapterStyle is not null || pageNumbering.ChapterSeparator is not null)
            return false;
        try
        {
            var value = new DocumentSectionPageNumbering();
            if (pageNumbering.Start?.Value is { } start)
            {
                if (start < 0) return false;
                value.Start = checked((uint)start);
            }
            if (pageNumbering.Format?.Value is { } format)
            {
                if (!TryFromNativePageNumberFormat(format, out var publicFormat)) return false;
                value.Format = publicFormat;
            }
            if (!value.HasStart && value.Format == DocumentSectionPageNumberFormat.Unspecified) return false;
            result = value;
            return true;
        }
        catch (Exception exception) when (exception is FormatException or InvalidOperationException or OverflowException)
        {
            return false;
        }
    }

    private static bool TryReadLineNumbering(W.SectionProperties source, out DocumentSectionLineNumbering? result)
    {
        result = null;
        var matches = source.Elements<W.LineNumberType>().ToArray();
        if (matches.Length == 0) return true;
        if (matches.Length != 1) return false;
        var lineNumbering = matches[0];
        if (lineNumbering.HasChildren || lineNumbering.ExtendedAttributes.Any() || lineNumbering.NamespaceDeclarations.Any() ||
            lineNumbering.MCAttributes is not null)
            return false;
        try
        {
            var countBy = lineNumbering.CountBy?.Value ?? (short)1;
            if (countBy is < 1 or > 32767) return false;
            var value = new DocumentSectionLineNumbering { CountBy = checked((uint)countBy) };
            if (lineNumbering.Start?.Value is { } start)
            {
                if (start < 0) return false;
                value.Start = checked((uint)start);
            }
            if (lineNumbering.Distance?.Value is { } distanceText)
            {
                if (!uint.TryParse(distanceText, NumberStyles.None, CultureInfo.InvariantCulture, out var distance) || distance > 31680)
                    return false;
                value.DistanceTwips = distance;
            }
            if (lineNumbering.Restart?.Value is { } restart)
            {
                if (!TryFromNativeLineNumberRestart(restart, out var publicRestart)) return false;
                value.Restart = publicRestart;
            }
            result = value;
            return true;
        }
        catch (Exception exception) when (exception is FormatException or InvalidOperationException or OverflowException)
        {
            return false;
        }
    }

    private static bool IsSupportedPageNumberFormat(DocumentSectionPageNumberFormat value) =>
        value is DocumentSectionPageNumberFormat.Unspecified or
            DocumentSectionPageNumberFormat.Decimal or
            DocumentSectionPageNumberFormat.UpperRoman or
            DocumentSectionPageNumberFormat.LowerRoman or
            DocumentSectionPageNumberFormat.UpperLetter or
            DocumentSectionPageNumberFormat.LowerLetter;

    private static W.NumberFormatValues ToNativePageNumberFormat(DocumentSectionPageNumberFormat value) => value switch
    {
        DocumentSectionPageNumberFormat.Decimal => W.NumberFormatValues.Decimal,
        DocumentSectionPageNumberFormat.UpperRoman => W.NumberFormatValues.UpperRoman,
        DocumentSectionPageNumberFormat.LowerRoman => W.NumberFormatValues.LowerRoman,
        DocumentSectionPageNumberFormat.UpperLetter => W.NumberFormatValues.UpperLetter,
        DocumentSectionPageNumberFormat.LowerLetter => W.NumberFormatValues.LowerLetter,
        _ => throw new CodecException("invalid_document_section", "Document section page-number format is unsupported."),
    };

    private static bool TryFromNativePageNumberFormat(W.NumberFormatValues value, out DocumentSectionPageNumberFormat result)
    {
        result = value == W.NumberFormatValues.Decimal ? DocumentSectionPageNumberFormat.Decimal :
            value == W.NumberFormatValues.UpperRoman ? DocumentSectionPageNumberFormat.UpperRoman :
            value == W.NumberFormatValues.LowerRoman ? DocumentSectionPageNumberFormat.LowerRoman :
            value == W.NumberFormatValues.UpperLetter ? DocumentSectionPageNumberFormat.UpperLetter :
            value == W.NumberFormatValues.LowerLetter ? DocumentSectionPageNumberFormat.LowerLetter :
            DocumentSectionPageNumberFormat.Unspecified;
        return result != DocumentSectionPageNumberFormat.Unspecified;
    }

    private static bool IsSupportedLineNumberRestart(DocumentSectionLineNumberRestart value) =>
        value is DocumentSectionLineNumberRestart.Unspecified or
            DocumentSectionLineNumberRestart.NewPage or
            DocumentSectionLineNumberRestart.NewSection or
            DocumentSectionLineNumberRestart.Continuous;

    private static W.LineNumberRestartValues ToNativeLineNumberRestart(DocumentSectionLineNumberRestart value) => value switch
    {
        DocumentSectionLineNumberRestart.NewPage => W.LineNumberRestartValues.NewPage,
        DocumentSectionLineNumberRestart.NewSection => W.LineNumberRestartValues.NewSection,
        DocumentSectionLineNumberRestart.Continuous => W.LineNumberRestartValues.Continuous,
        _ => throw new CodecException("invalid_document_section", "Document section line-number restart is unsupported."),
    };

    private static bool TryFromNativeLineNumberRestart(W.LineNumberRestartValues value, out DocumentSectionLineNumberRestart result)
    {
        result = value == W.LineNumberRestartValues.NewPage ? DocumentSectionLineNumberRestart.NewPage :
            value == W.LineNumberRestartValues.NewSection ? DocumentSectionLineNumberRestart.NewSection :
            value == W.LineNumberRestartValues.Continuous ? DocumentSectionLineNumberRestart.Continuous :
            DocumentSectionLineNumberRestart.Unspecified;
        return result != DocumentSectionLineNumberRestart.Unspecified;
    }

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
            var rank = SectionChildRank(replacement);
            var anchor = owner.ChildElements.FirstOrDefault(child => SectionChildRank(child) > rank);
            if (anchor is null) owner.Append(replacement);
            else owner.InsertBefore(replacement, anchor);
        }
        else current.InsertAfterSelf(replacement);
        current?.Remove();
    }

    private static int SectionChildRank(OpenXmlElement value) => value switch
    {
        W.HeaderReference => 10,
        W.FooterReference => 20,
        W.SectionType => 30,
        W.PageSize => 40,
        W.PageMargin => 50,
        W.LineNumberType => 60,
        W.PageNumberType => 70,
        W.Columns => 80,
        W.TitlePage => 90,
        _ => int.MaxValue,
    };

    private static string Hash(string value) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();
}
