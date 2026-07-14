using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OpenChestnut.Codec;

// Owns the modeled a:defRPr subset. Unknown attributes and children remain in
// the source element; explicit deletion clears only fields represented here.
internal static class PptxDefaultRunStyleCodec
{
    private const double MaxFontSizePoints = 768;

    internal static void Read(PresentationTextParagraph target, A.TextParagraphPropertiesType? source)
    {
        var properties = source?.GetFirstChild<A.DefaultRunProperties>();
        if (properties is null) return;
        var style = new PresentationTextStyle();
        if (properties.Bold is not null) style.Bold = properties.Bold.Value;
        if (properties.Italic is not null) style.Italic = properties.Italic.Value;
        if (properties.FontSize is not null) style.FontSizePoints = properties.FontSize.Value / 100d;
        var latin = properties.Elements<A.LatinFont>().SingleOrDefault();
        if (latin is not null && ModeledLatinFont(latin)) style.FontFamily = latin.Typeface!.Value!;
        var colors = ColorChoices(properties).ToArray();
        if (colors.Length == 1 && ModeledColor(colors[0]))
        {
            var fill = (A.SolidFill)colors[0];
            if (fill.GetFirstChild<A.RgbColorModelHex>() is { } rgb)
                style.ColorRgb = PptxColor.Normalize(rgb.Val!.Value!);
            else if (fill.GetFirstChild<A.SchemeColor>() is { } scheme && PptxColor.TrySchemeToken(scheme.Val!.Value, out var token))
                style.ColorScheme = token;
        }
        if (HasFields(style)) target.DefaultRunProperties = style;
    }

    internal static bool Supports(A.TextParagraphPropertiesType? source) =>
        source?.Elements<A.DefaultRunProperties>().Take(2).Count() is not > 1;

    internal static void Validate(PresentationTextParagraph paragraph)
    {
        switch (paragraph.DefaultRunStyleCase)
        {
            case PresentationTextParagraph.DefaultRunStyleOneofCase.None:
                return;
            case PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties:
                ValidateStyle(paragraph.DefaultRunProperties);
                return;
            case PresentationTextParagraph.DefaultRunStyleOneofCase.NoDefaultRunProperties:
                if (!paragraph.NoDefaultRunProperties)
                    throw Invalid("Presentation no_default_run_properties must be true when selected.");
                return;
            default:
                throw Invalid("Presentation paragraph contains an unknown default-run-style case.");
        }
    }

    internal static bool HasAuthoredStyle(PresentationTextParagraph paragraph) =>
        paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties;

    internal static void Append(A.TextParagraphPropertiesType target, PresentationTextParagraph source)
    {
        if (!HasAuthoredStyle(source)) return;
        target.AddChild(Build(source.DefaultRunProperties), true);
    }

    internal static void Apply(A.TextParagraphPropertiesType target, PresentationTextParagraph source)
    {
        if (source.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.None) return;
        var properties = target.GetFirstChild<A.DefaultRunProperties>();
        if (source.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.NoDefaultRunProperties)
        {
            if (properties is null) return;
            ClearModeled(properties);
            RemoveIfEmpty(properties);
            return;
        }
        if (properties is null)
        {
            properties = new A.DefaultRunProperties();
            target.AddChild(properties, true);
        }
        ApplyStyle(properties, source.DefaultRunProperties);
    }

    internal static void Scrub(A.TextParagraphPropertiesType target)
    {
        var properties = target.GetFirstChild<A.DefaultRunProperties>();
        if (properties is null) return;
        ClearModeled(properties);
        RemoveIfEmpty(properties);
    }

    private static void ValidateStyle(PresentationTextStyle style)
    {
        if (!HasFields(style)) throw Invalid("Presentation default run properties must contain at least one modeled field.");
        if (style.HasFontSizePoints && (!(style.FontSizePoints > 0) || style.FontSizePoints > MaxFontSizePoints || !double.IsFinite(style.FontSizePoints)))
            throw Invalid($"Presentation default-run font size must be finite and between 0 and {MaxFontSizePoints} points.");
        if (style.HasFontFamily && (string.IsNullOrWhiteSpace(style.FontFamily) || style.FontFamily.Length > 255))
            throw Invalid("Presentation default-run font family must contain 1 through 255 characters.");
        switch (style.ColorCase)
        {
            case PresentationTextStyle.ColorOneofCase.None:
                break;
            case PresentationTextStyle.ColorOneofCase.ColorRgb:
                _ = PptxColor.Normalize(style.ColorRgb);
                break;
            case PresentationTextStyle.ColorOneofCase.ColorScheme:
                _ = PptxColor.NormalizeScheme(style.ColorScheme);
                break;
            default:
                throw Invalid("Presentation default run properties contain an unknown color case.");
        }
    }

    private static bool HasFields(PresentationTextStyle style) =>
        style.HasBold || style.HasItalic || style.HasFontSizePoints || style.HasFontFamily ||
        style.ColorCase != PresentationTextStyle.ColorOneofCase.None;

    private static A.DefaultRunProperties Build(PresentationTextStyle source)
    {
        var target = new A.DefaultRunProperties();
        ApplyStyle(target, source);
        return target;
    }

    private static void ApplyStyle(A.DefaultRunProperties target, PresentationTextStyle source)
    {
        target.Bold = source.HasBold ? source.Bold : null;
        target.Italic = source.HasItalic ? source.Italic : null;
        target.FontSize = source.HasFontSizePoints ? checked((int)Math.Round(source.FontSizePoints * 100)) : null;
        ApplyLatinFont(target, source);
        ApplyColor(target, source);
    }

    private static void ApplyLatinFont(A.DefaultRunProperties target, PresentationTextStyle source)
    {
        var fonts = target.Elements<A.LatinFont>().ToArray();
        if (source.HasFontFamily)
        {
            if (fonts.Length > 1 || fonts.Any(font => !ModeledLatinFont(font)))
                throw Unsupported("Source-preserving PPTX export cannot replace unmodeled default-run Latin font properties.");
            foreach (var font in fonts) font.Remove();
            target.AddChild(new A.LatinFont { Typeface = source.FontFamily }, true);
        }
        else if (fonts.Length == 1 && ModeledLatinFont(fonts[0]))
        {
            fonts[0].Remove();
        }
    }

    private static void ApplyColor(A.DefaultRunProperties target, PresentationTextStyle source)
    {
        var colors = ColorChoices(target).ToArray();
        if (source.ColorCase != PresentationTextStyle.ColorOneofCase.None)
        {
            if (colors.Length > 1 || colors.Any(color => !ModeledColor(color)))
                throw Unsupported("Source-preserving PPTX export cannot replace unmodeled default-run color properties.");
            foreach (var color in colors) color.Remove();
            target.AddChild(source.ColorCase == PresentationTextStyle.ColorOneofCase.ColorRgb
                ? new A.SolidFill(new A.RgbColorModelHex { Val = PptxColor.Normalize(source.ColorRgb) })
                : new A.SolidFill(new A.SchemeColor { Val = PptxColor.SchemeValue(source.ColorScheme) }), true);
        }
        else if (colors.Length == 1 && ModeledColor(colors[0]))
        {
            colors[0].Remove();
        }
    }

    private static void ClearModeled(A.DefaultRunProperties target)
    {
        target.Bold = null;
        target.Italic = null;
        target.FontSize = null;
        var fonts = target.Elements<A.LatinFont>().ToArray();
        if (fonts.Length == 1 && ModeledLatinFont(fonts[0])) fonts[0].Remove();
        var colors = ColorChoices(target).ToArray();
        if (colors.Length == 1 && ModeledColor(colors[0])) colors[0].Remove();
    }

    private static IEnumerable<OpenXmlElement> ColorChoices(A.DefaultRunProperties source) =>
        source.ChildElements.Where(child => child.LocalName is "noFill" or "solidFill" or "gradFill" or "blipFill" or "pattFill" or "grpFill");

    private static bool ModeledColor(OpenXmlElement source) => source switch
    {
        A.SolidFill fill when fill.ChildElements.Count == 1 && fill.GetFirstChild<A.RgbColorModelHex>() is { } rgb =>
            SimpleValue(rgb, "val") && ValidRgb(rgb.Val?.Value),
        A.SolidFill fill when fill.ChildElements.Count == 1 && fill.GetFirstChild<A.SchemeColor>() is { } scheme =>
            SimpleValue(scheme, "val") && scheme.Val?.Value is { } value && PptxColor.TrySchemeToken(value, out _),
        _ => false,
    };

    private static bool ModeledLatinFont(A.LatinFont source) =>
        SimpleValue(source, "typeface") && !string.IsNullOrWhiteSpace(source.Typeface?.Value) && source.Typeface.Value.Length <= 255;

    private static bool SimpleValue(OpenXmlElement source, string name)
    {
        var attributes = source.GetAttributes();
        return source.ChildElements.Count == 0 && attributes.Count == 1 && attributes[0].LocalName == name;
    }

    private static bool ValidRgb(string? value)
    {
        try
        {
            _ = PptxColor.Normalize(value ?? string.Empty);
            return true;
        }
        catch (CodecException)
        {
            return false;
        }
    }

    private static void RemoveIfEmpty(A.DefaultRunProperties target)
    {
        if (target.GetAttributes().Count == 0 && target.ChildElements.Count == 0) target.Remove();
    }

    private static CodecException Invalid(string message) => new("invalid_presentation_text", message);

    private static CodecException Unsupported(string message) => new("unsupported_presentation_edit", message);
}
