using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OpenOffice.OpenXmlCodec;

// Owns the three independent direct DrawingML marker-style choices. Unset wire
// choices leave unknown or inherited source styling untouched.
internal static class PptxBulletStyleCodec
{
    private const double MaxSizePoints = 768;

    internal static void Read(PresentationTextParagraph target, A.ParagraphProperties? source)
    {
        if (source is null) return;
        var font = FontChoices(source).ToArray();
        if (font.Length == 1 && ModeledFont(font[0]))
        {
            if (font[0] is A.BulletFont specified) target.BulletFontFamily = specified.Typeface!.Value!;
            else target.BulletFontFollowText = true;
        }

        var color = ColorChoices(source).ToArray();
        if (color.Length == 1 && ModeledColor(color[0]))
        {
            if (color[0] is A.BulletColor specified)
                target.BulletColorRgb = PptxColor.Normalize(specified.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value!);
            else target.BulletColorFollowText = true;
        }

        var size = SizeChoices(source).ToArray();
        if (size.Length == 1 && ModeledSize(size[0]))
        {
            switch (size[0])
            {
                case A.BulletSizePoints points:
                    target.BulletSizePoints = points.Val!.Value / 100d;
                    break;
                case A.BulletSizePercentage percent:
                    target.BulletSizePercent = percent.Val!.Value / 100_000d;
                    break;
                default:
                    target.BulletSizeFollowText = true;
                    break;
            }
        }
    }

    internal static void Validate(PresentationTextParagraph paragraph)
    {
        switch (paragraph.BulletFontCase)
        {
            case PresentationTextParagraph.BulletFontOneofCase.None:
                break;
            case PresentationTextParagraph.BulletFontOneofCase.BulletFontFamily:
                if (string.IsNullOrWhiteSpace(paragraph.BulletFontFamily) || paragraph.BulletFontFamily.Length > 255)
                    throw Invalid("Presentation bullet font family must contain 1 through 255 characters.");
                break;
            case PresentationTextParagraph.BulletFontOneofCase.BulletFontFollowText:
                if (!paragraph.BulletFontFollowText) throw Invalid("Presentation bullet_font_follow_text must be true when selected.");
                break;
            default:
                throw Invalid("Presentation paragraph contains an unknown bullet-font case.");
        }

        switch (paragraph.BulletColorCase)
        {
            case PresentationTextParagraph.BulletColorOneofCase.None:
                break;
            case PresentationTextParagraph.BulletColorOneofCase.BulletColorRgb:
                _ = PptxColor.Normalize(paragraph.BulletColorRgb);
                break;
            case PresentationTextParagraph.BulletColorOneofCase.BulletColorFollowText:
                if (!paragraph.BulletColorFollowText) throw Invalid("Presentation bullet_color_follow_text must be true when selected.");
                break;
            default:
                throw Invalid("Presentation paragraph contains an unknown bullet-color case.");
        }

        switch (paragraph.BulletSizeCase)
        {
            case PresentationTextParagraph.BulletSizeOneofCase.None:
                break;
            case PresentationTextParagraph.BulletSizeOneofCase.BulletSizePoints:
                if (!double.IsFinite(paragraph.BulletSizePoints) || paragraph.BulletSizePoints < 1 || paragraph.BulletSizePoints > MaxSizePoints)
                    throw Invalid($"Presentation bullet size must be from 1 through {MaxSizePoints} points.");
                break;
            case PresentationTextParagraph.BulletSizeOneofCase.BulletSizePercent:
                if (!double.IsFinite(paragraph.BulletSizePercent) || paragraph.BulletSizePercent < 0.25 || paragraph.BulletSizePercent > 4)
                    throw Invalid("Presentation bullet size percentage must be from 0.25 through 4.");
                break;
            case PresentationTextParagraph.BulletSizeOneofCase.BulletSizeFollowText:
                if (!paragraph.BulletSizeFollowText) throw Invalid("Presentation bullet_size_follow_text must be true when selected.");
                break;
            default:
                throw Invalid("Presentation paragraph contains an unknown bullet-size case.");
        }
    }

    internal static bool HasModeledStyle(PresentationTextParagraph paragraph) =>
        paragraph.BulletFontCase != PresentationTextParagraph.BulletFontOneofCase.None ||
        paragraph.BulletColorCase != PresentationTextParagraph.BulletColorOneofCase.None ||
        paragraph.BulletSizeCase != PresentationTextParagraph.BulletSizeOneofCase.None;

    internal static void Append(A.ParagraphProperties target, PresentationTextParagraph source)
    {
        if (source.BulletColorCase != PresentationTextParagraph.BulletColorOneofCase.None) target.AddChild(BuildColor(source), true);
        if (source.BulletSizeCase != PresentationTextParagraph.BulletSizeOneofCase.None) target.AddChild(BuildSize(source), true);
        if (source.BulletFontCase != PresentationTextParagraph.BulletFontOneofCase.None) target.AddChild(BuildFont(source), true);
    }

    internal static void Apply(A.ParagraphProperties target, PresentationTextParagraph source)
    {
        ApplyChoice(target, source.BulletColorCase != PresentationTextParagraph.BulletColorOneofCase.None, ColorChoices, ModeledColor, () => BuildColor(source), "color");
        ApplyChoice(target, source.BulletSizeCase != PresentationTextParagraph.BulletSizeOneofCase.None, SizeChoices, ModeledSize, () => BuildSize(source), "size");
        ApplyChoice(target, source.BulletFontCase != PresentationTextParagraph.BulletFontOneofCase.None, FontChoices, ModeledFont, () => BuildFont(source), "font");
    }

    internal static void Scrub(A.ParagraphProperties target)
    {
        ScrubChoice(target, ColorChoices, ModeledColor);
        ScrubChoice(target, SizeChoices, ModeledSize);
        ScrubChoice(target, FontChoices, ModeledFont);
    }

    private static void ApplyChoice(
        A.ParagraphProperties target,
        bool requested,
        Func<A.ParagraphProperties, IEnumerable<OpenXmlElement>> choices,
        Func<OpenXmlElement, bool> modeled,
        Func<OpenXmlElement> build,
        string kind)
    {
        if (!requested) return;
        var existing = choices(target).ToArray();
        if (existing.Length > 1 || existing.Any(choice => !modeled(choice)))
            throw new CodecException("unsupported_presentation_edit", $"Source-preserving PPTX export cannot replace an unmodeled or malformed bullet {kind}.");
        foreach (var choice in existing) choice.Remove();
        target.AddChild(build(), true);
    }

    private static void ScrubChoice(
        A.ParagraphProperties target,
        Func<A.ParagraphProperties, IEnumerable<OpenXmlElement>> choices,
        Func<OpenXmlElement, bool> modeled)
    {
        var existing = choices(target).ToArray();
        if (existing.Length == 1 && modeled(existing[0])) existing[0].Remove();
    }

    private static OpenXmlElement BuildFont(PresentationTextParagraph source) => source.BulletFontCase switch
    {
        PresentationTextParagraph.BulletFontOneofCase.BulletFontFamily => new A.BulletFont { Typeface = source.BulletFontFamily },
        PresentationTextParagraph.BulletFontOneofCase.BulletFontFollowText => new A.BulletFontText(),
        _ => throw Invalid("Presentation paragraph has no modeled bullet-font style."),
    };

    private static OpenXmlElement BuildColor(PresentationTextParagraph source) => source.BulletColorCase switch
    {
        PresentationTextParagraph.BulletColorOneofCase.BulletColorRgb => new A.BulletColor(new A.RgbColorModelHex { Val = PptxColor.Normalize(source.BulletColorRgb) }),
        PresentationTextParagraph.BulletColorOneofCase.BulletColorFollowText => new A.BulletColorText(),
        _ => throw Invalid("Presentation paragraph has no modeled bullet-color style."),
    };

    private static OpenXmlElement BuildSize(PresentationTextParagraph source) => source.BulletSizeCase switch
    {
        PresentationTextParagraph.BulletSizeOneofCase.BulletSizePoints => new A.BulletSizePoints { Val = checked((int)Math.Round(source.BulletSizePoints * 100)) },
        PresentationTextParagraph.BulletSizeOneofCase.BulletSizePercent => new A.BulletSizePercentage { Val = checked((int)Math.Round(source.BulletSizePercent * 100_000)) },
        PresentationTextParagraph.BulletSizeOneofCase.BulletSizeFollowText => new A.BulletSizeText(),
        _ => throw Invalid("Presentation paragraph has no modeled bullet-size style."),
    };

    private static IEnumerable<OpenXmlElement> FontChoices(A.ParagraphProperties source) =>
        source.ChildElements.Where(child => child is A.BulletFont or A.BulletFontText);

    private static IEnumerable<OpenXmlElement> ColorChoices(A.ParagraphProperties source) =>
        source.ChildElements.Where(child => child is A.BulletColor or A.BulletColorText);

    private static IEnumerable<OpenXmlElement> SizeChoices(A.ParagraphProperties source) =>
        source.ChildElements.Where(child => child is A.BulletSizePoints or A.BulletSizePercentage or A.BulletSizeText);

    private static bool ModeledFont(OpenXmlElement source) => source switch
    {
        A.BulletFont font => SimpleAttribute(font, "typeface") && !string.IsNullOrWhiteSpace(font.Typeface?.Value) && font.Typeface.Value.Length <= 255,
        A.BulletFontText follow => Empty(follow),
        _ => false,
    };

    private static bool ModeledColor(OpenXmlElement source) => source switch
    {
        A.BulletColor color when EmptyAttributes(color) && color.ChildElements.Count == 1 && color.GetFirstChild<A.RgbColorModelHex>() is { } rgb =>
            SimpleAttribute(rgb, "val") && rgb.ChildElements.Count == 0 && ValidRgb(rgb.Val?.Value),
        A.BulletColorText follow => Empty(follow),
        _ => false,
    };

    private static bool ModeledSize(OpenXmlElement source) => source switch
    {
        A.BulletSizePoints points => SimpleAttribute(points, "val") && points.Val?.Value is >= 100 and <= 76_800,
        A.BulletSizePercentage percent => SimpleAttribute(percent, "val") && percent.Val?.Value is >= 25_000 and <= 400_000,
        A.BulletSizeText follow => Empty(follow),
        _ => false,
    };

    private static bool Empty(OpenXmlElement source) => EmptyAttributes(source) && source.ChildElements.Count == 0;

    private static bool EmptyAttributes(OpenXmlElement source) => source.GetAttributes().Count == 0;

    private static bool SimpleAttribute(OpenXmlElement source, string name)
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

    private static CodecException Invalid(string message) => new("invalid_presentation_text", message);
}
