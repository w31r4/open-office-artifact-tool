using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Owns the deliberately small direct-formatting profile shared by document
// defaults, styles, paragraphs, and ordinary runs. Theme fonts/colors and
// other inherited graphs remain source-owned.
internal static class DocxFormattingCodec
{
    private static readonly HashSet<string> Alignments = new(StringComparer.Ordinal)
    {
        "left", "center", "right", "justify",
    };

    private static readonly HashSet<string> LineRules = new(StringComparer.Ordinal)
    {
        "auto", "atLeast", "exact",
    };

    internal static DocumentRunFormatting? ReadRunFormatting(W.RunProperties? properties)
    {
        if (properties is null) return null;
        if (properties.ChildElements.Any(child => child is not W.RunStyle and
                                                   not W.RunFonts and
                                                   not W.Bold and
                                                   not W.Italic and
                                                   not W.Underline and
                                                   not W.FontSize and
                                                   not W.Color and
                                                   not W.Spacing)) return null;

        var fonts = properties.RunFonts;
        if (fonts is not null &&
            (fonts.AsciiTheme is not null || fonts.HighAnsiTheme is not null ||
             fonts.EastAsiaTheme is not null || fonts.ComplexScriptTheme is not null ||
             fonts.EastAsia is not null || fonts.ComplexScript is not null ||
             (fonts.Ascii?.Value is { } ascii && fonts.HighAnsi?.Value is { } highAnsi &&
              !ascii.Equals(highAnsi, StringComparison.Ordinal)))) return null;

        var color = properties.Color;
        if (color is not null &&
            (color.ThemeColor is not null || color.ThemeTint is not null || color.ThemeShade is not null ||
             !IsRgb(color.Val?.Value))) return null;

        if (!TryHalfPoints(properties.FontSize?.Val?.Value, out var halfPoints)) return null;
        int? characterSpacing = properties.Spacing?.Val?.Value;

        var result = new DocumentRunFormatting();
        var fontFamily = fonts?.Ascii?.Value ?? fonts?.HighAnsi?.Value;
        if (!string.IsNullOrWhiteSpace(fontFamily)) result.FontFamily = fontFamily;
        if (halfPoints is not null) result.FontSizeHalfPoints = halfPoints.Value;
        if (color?.Val?.Value is { } rgb) result.ColorRgb = rgb.ToUpperInvariant();
        if (characterSpacing is not null) result.CharacterSpacingTwips = characterSpacing.Value;
        if (properties.Bold is { } bold) result.Bold = IsOn(bold);
        if (properties.Italic is { } italic) result.Italic = IsOn(italic);
        if (properties.Underline is { } underline)
            result.Underline = underline.Val?.Value is not null && underline.Val.Value != W.UnderlineValues.None;
        return HasRunFormatting(result) ? result : null;
    }

    internal static DocumentRunFormatting? ReadStyleRunFormatting(W.StyleRunProperties? properties)
    {
        if (properties is null) return null;
        var probe = new W.RunProperties();
        foreach (var child in properties.ChildElements) probe.Append(child.CloneNode(true));
        return ReadRunFormatting(probe);
    }

    internal static DocumentRunFormatting? ReadBaseRunFormatting(W.RunPropertiesBaseStyle? properties)
    {
        if (properties is null) return null;
        var probe = new W.RunProperties();
        foreach (var child in properties.ChildElements) probe.Append(child.CloneNode(true));
        return ReadRunFormatting(probe);
    }

    internal static DocumentParagraphFormatting? ReadParagraphFormatting(W.ParagraphProperties? properties)
    {
        if (properties is null) return null;
        var result = ReadParagraphFormattingCore(properties.Justification, properties.Indentation,
            properties.SpacingBetweenLines, properties.KeepNext, properties.PageBreakBefore);
        return HasParagraphFormatting(result) ? result : null;
    }

    internal static DocumentParagraphFormatting? ReadStyleParagraphFormatting(W.StyleParagraphProperties? properties)
    {
        if (properties is null) return null;
        var result = ReadParagraphFormattingCore(properties.Justification, properties.Indentation,
            properties.SpacingBetweenLines, properties.KeepNext, properties.PageBreakBefore);
        return HasParagraphFormatting(result) ? result : null;
    }

    internal static W.RunProperties? BuildRunProperties(DocumentRun source)
    {
        var formatting = MergeLegacy(source);
        Validate(formatting, $"Document run in {source.Text}");
        var properties = new W.RunProperties();
        if (!string.IsNullOrWhiteSpace(source.StyleId)) properties.Append(new W.RunStyle { Val = source.StyleId });
        AppendRunFormatting(properties, formatting);
        return properties.ChildElements.Count == 0 ? null : properties;
    }

    internal static W.StyleRunProperties? BuildStyleRunProperties(DocumentRunFormatting? formatting, string label)
    {
        Validate(formatting, label);
        if (!HasRunFormatting(formatting)) return null;
        var properties = new W.StyleRunProperties();
        AppendRunFormatting(properties, formatting!);
        return properties;
    }

    internal static W.RunPropertiesBaseStyle? BuildDefaultRunProperties(DocumentRunFormatting? formatting)
    {
        Validate(formatting, "Document default run style");
        if (!HasRunFormatting(formatting)) return null;
        var properties = new W.RunPropertiesBaseStyle();
        AppendRunFormatting(properties, formatting!);
        return properties;
    }

    internal static W.ParagraphProperties? BuildParagraphProperties(
        string styleId,
        DocumentParagraphFormatting? formatting,
        DocumentNumbering? numbering = null)
    {
        Validate(formatting, "Document paragraph formatting");
        var properties = new W.ParagraphProperties();
        if (!string.IsNullOrWhiteSpace(styleId)) properties.Append(new W.ParagraphStyleId { Val = styleId });
        AppendParagraphFormattingBeforeNumbering(properties, formatting);
        if (numbering is not null)
            properties.Append(new W.NumberingProperties(
                new W.NumberingLevelReference { Val = checked((int)numbering.Level) },
                new W.NumberingId { Val = checked((int)numbering.NumberingId) }));
        AppendParagraphFormattingAfterNumbering(properties, formatting);
        return properties.ChildElements.Count == 0 ? null : properties;
    }

    internal static W.StyleParagraphProperties? BuildStyleParagraphProperties(DocumentParagraphFormatting? formatting, string label)
    {
        Validate(formatting, label);
        if (!HasParagraphFormatting(formatting)) return null;
        var properties = new W.StyleParagraphProperties();
        AppendParagraphFormattingBeforeNumbering(properties, formatting);
        AppendParagraphFormattingAfterNumbering(properties, formatting);
        return properties;
    }

    internal static bool IsSupportedParagraphProperties(W.ParagraphProperties? properties, bool allowNumbering = false, bool allowSection = false)
    {
        if (properties is null) return true;
        return properties.ChildElements.All(child => child is W.ParagraphStyleId or W.Justification or W.Indentation or
            W.SpacingBetweenLines or W.KeepNext or W.PageBreakBefore ||
            (allowNumbering && child is W.NumberingProperties) ||
            (allowSection && child is W.SectionProperties));
    }

    internal static bool IsSupportedRunProperties(W.RunProperties? properties) =>
        properties is null || ReadRunFormatting(properties) is not null ||
        properties.ChildElements.All(child => child is W.RunStyle);

    internal static void Validate(DocumentRunFormatting? formatting, string label)
    {
        if (formatting is null) return;
        if (formatting.HasFontFamily &&
            (string.IsNullOrWhiteSpace(formatting.FontFamily) || formatting.FontFamily.Length > 255 || formatting.FontFamily.Any(char.IsControl)))
            throw new CodecException("invalid_document_run_formatting", $"{label} font family must contain 1 through 255 characters without controls.");
        if (formatting.HasFontSizeHalfPoints && formatting.FontSizeHalfPoints is < 1 or > 3276)
            throw new CodecException("invalid_document_run_formatting", $"{label} font size must be 1 through 3276 half-points.");
        if (formatting.HasColorRgb && !IsRgb(formatting.ColorRgb))
            throw new CodecException("invalid_document_run_formatting", $"{label} color must be a six-digit RGB value.");
        if (formatting.HasCharacterSpacingTwips && formatting.CharacterSpacingTwips is < -31680 or > 31680)
            throw new CodecException("invalid_document_run_formatting", $"{label} character spacing must be between -31680 and 31680 twentieths of a point.");
    }

    internal static void Validate(DocumentParagraphFormatting? formatting, string label)
    {
        if (formatting is null) return;
        if (formatting.HasAlignment && !Alignments.Contains(formatting.Alignment))
            throw new CodecException("invalid_document_paragraph_formatting", $"{label} alignment must be left, center, right, or justify.");
        if (formatting.HasFirstLineIndentTwips && formatting.HasHangingIndentTwips)
            throw new CodecException("invalid_document_paragraph_formatting", $"{label} cannot set first-line and hanging indent together.");
        foreach (var (present, value, name) in new[]
                 {
                     (formatting.HasSpaceBeforeTwips, formatting.SpaceBeforeTwips, "space before"),
                     (formatting.HasSpaceAfterTwips, formatting.SpaceAfterTwips, "space after"),
                     (formatting.HasLineSpacingTwips, formatting.LineSpacingTwips, "line spacing"),
                 })
            if (present && value is < 0 or > 31680)
                throw new CodecException("invalid_document_paragraph_formatting", $"{label} {name} must be between 0 and 31680 twentieths of a point.");
        foreach (var (present, value, name) in new[]
                 {
                     (formatting.HasLeftIndentTwips, formatting.LeftIndentTwips, "left indent"),
                     (formatting.HasRightIndentTwips, formatting.RightIndentTwips, "right indent"),
                     (formatting.HasFirstLineIndentTwips, formatting.FirstLineIndentTwips, "first-line indent"),
                     (formatting.HasHangingIndentTwips, formatting.HangingIndentTwips, "hanging indent"),
                 })
            if (present && value is < -31680 or > 31680)
                throw new CodecException("invalid_document_paragraph_formatting", $"{label} {name} must be between -31680 and 31680 twentieths of a point.");
        if (formatting.HasLineSpacingRule && !LineRules.Contains(formatting.LineSpacingRule))
            throw new CodecException("invalid_document_paragraph_formatting", $"{label} line-spacing rule must be auto, atLeast, or exact.");
        if (formatting.HasLineSpacingRule && !formatting.HasLineSpacingTwips)
            throw new CodecException("invalid_document_paragraph_formatting", $"{label} line-spacing rule requires line spacing.");
    }

    internal static DocumentRunFormatting MergeLegacy(DocumentRun source)
    {
        var result = source.Formatting?.Clone() ?? new DocumentRunFormatting();
        if (!result.HasBold && source.Bold) result.Bold = true;
        if (!result.HasItalic && source.Italic) result.Italic = true;
        if (!result.HasUnderline && source.Underline) result.Underline = true;
        return result;
    }

    internal static bool HasRunFormatting(DocumentRunFormatting? value) => value is not null &&
        (value.HasFontFamily || value.HasFontSizeHalfPoints || value.HasColorRgb || value.HasCharacterSpacingTwips ||
         value.HasBold || value.HasItalic || value.HasUnderline);

    internal static bool HasParagraphFormatting(DocumentParagraphFormatting? value) => value is not null &&
        (value.HasAlignment || value.HasLeftIndentTwips || value.HasRightIndentTwips || value.HasFirstLineIndentTwips ||
         value.HasHangingIndentTwips || value.HasSpaceBeforeTwips || value.HasSpaceAfterTwips || value.HasLineSpacingTwips ||
         value.HasLineSpacingRule || value.HasKeepNext || value.HasPageBreakBefore);

    private static DocumentParagraphFormatting ReadParagraphFormattingCore(
        W.Justification? justification,
        W.Indentation? indentation,
        W.SpacingBetweenLines? spacing,
        W.KeepNext? keepNext,
        W.PageBreakBefore? pageBreakBefore)
    {
        var result = new DocumentParagraphFormatting();
        var nativeAlignment = justification?.Val?.Value;
        var alignment = nativeAlignment == W.JustificationValues.Left || nativeAlignment == W.JustificationValues.Start
            ? "left"
            : nativeAlignment == W.JustificationValues.Center
                ? "center"
                : nativeAlignment == W.JustificationValues.Right || nativeAlignment == W.JustificationValues.End
                    ? "right"
                    : nativeAlignment == W.JustificationValues.Both ? "justify" : null;
        if (alignment is not null) result.Alignment = alignment;
        if (TryInt(indentation?.Left?.Value ?? indentation?.Start?.Value, out var left) && left is not null) result.LeftIndentTwips = left.Value;
        if (TryInt(indentation?.Right?.Value ?? indentation?.End?.Value, out var right) && right is not null) result.RightIndentTwips = right.Value;
        if (TryInt(indentation?.FirstLine?.Value, out var firstLine) && firstLine is not null) result.FirstLineIndentTwips = firstLine.Value;
        if (TryInt(indentation?.Hanging?.Value, out var hanging) && hanging is not null) result.HangingIndentTwips = hanging.Value;
        if (TryInt(spacing?.Before?.Value, out var before) && before is not null) result.SpaceBeforeTwips = before.Value;
        if (TryInt(spacing?.After?.Value, out var after) && after is not null) result.SpaceAfterTwips = after.Value;
        var hasLineSpacing = TryInt(spacing?.Line?.Value, out var line) && line is not null;
        if (hasLineSpacing) result.LineSpacingTwips = line!.Value;
        var nativeRule = spacing?.LineRule?.Value;
        var rule = nativeRule == W.LineSpacingRuleValues.Auto
            ? "auto"
            : nativeRule == W.LineSpacingRuleValues.AtLeast
                ? "atLeast"
                : nativeRule == W.LineSpacingRuleValues.Exact ? "exact" : null;
        // Word may emit lineRule="auto" alongside before/after spacing while
        // omitting w:line. That attribute has no modeled line-height value and
        // cannot round-trip through the source-free authoring contract (which
        // requires a line value when a rule is explicit). Normalize it away;
        // source-bound style catalogs remain byte-preserved and edits are
        // rejected by DocxDirectStyles.AssertSourceUnchanged.
        if (hasLineSpacing && rule is not null) result.LineSpacingRule = rule;
        if (keepNext is not null) result.KeepNext = IsOn(keepNext);
        if (pageBreakBefore is not null) result.PageBreakBefore = IsOn(pageBreakBefore);
        return result;
    }

    private static void AppendRunFormatting(OpenXmlCompositeElement properties, DocumentRunFormatting formatting)
    {
        if (formatting.HasFontFamily)
            properties.Append(new W.RunFonts { Ascii = formatting.FontFamily, HighAnsi = formatting.FontFamily });
        if (formatting.HasBold) properties.Append(new W.Bold { Val = formatting.Bold });
        if (formatting.HasItalic) properties.Append(new W.Italic { Val = formatting.Italic });
        if (formatting.HasColorRgb) properties.Append(new W.Color { Val = formatting.ColorRgb.ToUpperInvariant() });
        if (formatting.HasCharacterSpacingTwips) properties.Append(new W.Spacing { Val = formatting.CharacterSpacingTwips });
        if (formatting.HasFontSizeHalfPoints) properties.Append(new W.FontSize { Val = formatting.FontSizeHalfPoints.ToString() });
        if (formatting.HasUnderline)
            properties.Append(new W.Underline { Val = formatting.Underline ? W.UnderlineValues.Single : W.UnderlineValues.None });
    }

    private static void AppendParagraphFormattingBeforeNumbering(
        OpenXmlCompositeElement properties,
        DocumentParagraphFormatting? formatting)
    {
        if (formatting is null) return;
        if (formatting.HasKeepNext) properties.Append(new W.KeepNext { Val = formatting.KeepNext });
        if (formatting.HasPageBreakBefore) properties.Append(new W.PageBreakBefore { Val = formatting.PageBreakBefore });
    }

    private static void AppendParagraphFormattingAfterNumbering(
        OpenXmlCompositeElement properties,
        DocumentParagraphFormatting? formatting)
    {
        if (formatting is null) return;
        if (formatting.HasSpaceBeforeTwips || formatting.HasSpaceAfterTwips || formatting.HasLineSpacingTwips)
        {
            var spacing = new W.SpacingBetweenLines();
            if (formatting.HasSpaceBeforeTwips) spacing.Before = formatting.SpaceBeforeTwips.ToString();
            if (formatting.HasSpaceAfterTwips) spacing.After = formatting.SpaceAfterTwips.ToString();
            if (formatting.HasLineSpacingTwips) spacing.Line = formatting.LineSpacingTwips.ToString();
            if (formatting.HasLineSpacingRule)
                spacing.LineRule = formatting.LineSpacingRule switch
                {
                    "auto" => W.LineSpacingRuleValues.Auto,
                    "atLeast" => W.LineSpacingRuleValues.AtLeast,
                    _ => W.LineSpacingRuleValues.Exact,
                };
            properties.Append(spacing);
        }
        if (formatting.HasLeftIndentTwips || formatting.HasRightIndentTwips ||
            formatting.HasFirstLineIndentTwips || formatting.HasHangingIndentTwips)
        {
            var indentation = new W.Indentation();
            if (formatting.HasLeftIndentTwips) indentation.Left = formatting.LeftIndentTwips.ToString();
            if (formatting.HasRightIndentTwips) indentation.Right = formatting.RightIndentTwips.ToString();
            if (formatting.HasFirstLineIndentTwips) indentation.FirstLine = formatting.FirstLineIndentTwips.ToString();
            if (formatting.HasHangingIndentTwips) indentation.Hanging = formatting.HangingIndentTwips.ToString();
            properties.Append(indentation);
        }
        if (formatting.HasAlignment)
            properties.Append(new W.Justification
            {
                Val = formatting.Alignment switch
                {
                    "left" => W.JustificationValues.Left,
                    "center" => W.JustificationValues.Center,
                    "right" => W.JustificationValues.Right,
                    _ => W.JustificationValues.Both,
                },
            });
    }

    private static bool IsOn(W.OnOffType value) => value.Val?.Value != false;
    private static bool IsRgb(string? value) => value is { Length: 6 } && value.All(char.IsAsciiHexDigit);

    private static bool TryHalfPoints(string? value, out uint? result)
    {
        result = null;
        if (value is null) return true;
        if (!uint.TryParse(value, out var parsed) || parsed is < 1 or > 3276) return false;
        result = parsed;
        return true;
    }

    private static bool TryInt(string? value, out int? result)
    {
        result = null;
        if (value is null) return true;
        if (!int.TryParse(value, out var parsed)) return false;
        result = parsed;
        return true;
    }
}
