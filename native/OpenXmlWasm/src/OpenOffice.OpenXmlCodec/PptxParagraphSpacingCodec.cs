using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OpenOffice.OpenXmlCodec;

// Owns the three direct DrawingML paragraph-spacing slots. Point and percent
// children remain distinct so an import/edit/export cycle never guesses units.
internal static class PptxParagraphSpacingCodec
{
    private const int MaxPointsHundredths = 158_400;
    private const int MaxPercentThousandths = 13_200_000;

    internal static void Read(PresentationTextParagraph target, A.TextParagraphPropertiesType? source)
    {
        ReadSlot(source?.GetFirstChild<A.LineSpacing>(), false,
            points => target.LineSpacingPoints = points,
            multiplier => target.LineSpacingMultiplier = multiplier);
        ReadSlot(source?.GetFirstChild<A.SpaceBefore>(), true,
            points => target.SpaceBeforePoints = points,
            multiplier => target.SpaceBeforeMultiplier = multiplier);
        ReadSlot(source?.GetFirstChild<A.SpaceAfter>(), true,
            points => target.SpaceAfterPoints = points,
            multiplier => target.SpaceAfterMultiplier = multiplier);
    }

    internal static bool Supports(A.TextParagraphPropertiesType? source) =>
        source is null ||
        SupportsSingle(source.Elements<A.LineSpacing>(), false) &&
        SupportsSingle(source.Elements<A.SpaceBefore>(), true) &&
        SupportsSingle(source.Elements<A.SpaceAfter>(), true);

    internal static void Validate(PresentationTextParagraph source)
    {
        switch (source.LineSpacingCase)
        {
            case PresentationTextParagraph.LineSpacingOneofCase.None:
                break;
            case PresentationTextParagraph.LineSpacingOneofCase.LineSpacingPoints:
                _ = Points(source.LineSpacingPoints, false, "line spacing");
                break;
            case PresentationTextParagraph.LineSpacingOneofCase.LineSpacingMultiplier:
                _ = Percent(source.LineSpacingMultiplier, false, "line spacing");
                break;
            case PresentationTextParagraph.LineSpacingOneofCase.NoLineSpacing:
                if (!source.NoLineSpacing) throw Invalid("Presentation no_line_spacing must be true when selected.");
                break;
            default:
                throw Invalid("Presentation paragraph contains an unknown line-spacing case.");
        }

        switch (source.SpaceBeforeCase)
        {
            case PresentationTextParagraph.SpaceBeforeOneofCase.None:
                break;
            case PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforePoints:
                _ = Points(source.SpaceBeforePoints, true, "space before");
                break;
            case PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforeMultiplier:
                _ = Percent(source.SpaceBeforeMultiplier, true, "space before");
                break;
            case PresentationTextParagraph.SpaceBeforeOneofCase.NoSpaceBefore:
                if (!source.NoSpaceBefore) throw Invalid("Presentation no_space_before must be true when selected.");
                break;
            default:
                throw Invalid("Presentation paragraph contains an unknown space-before case.");
        }

        switch (source.SpaceAfterCase)
        {
            case PresentationTextParagraph.SpaceAfterOneofCase.None:
                break;
            case PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterPoints:
                _ = Points(source.SpaceAfterPoints, true, "space after");
                break;
            case PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterMultiplier:
                _ = Percent(source.SpaceAfterMultiplier, true, "space after");
                break;
            case PresentationTextParagraph.SpaceAfterOneofCase.NoSpaceAfter:
                if (!source.NoSpaceAfter) throw Invalid("Presentation no_space_after must be true when selected.");
                break;
            default:
                throw Invalid("Presentation paragraph contains an unknown space-after case.");
        }
    }

    internal static bool HasAuthoredSpacing(PresentationTextParagraph source) =>
        source.LineSpacingCase is PresentationTextParagraph.LineSpacingOneofCase.LineSpacingPoints or PresentationTextParagraph.LineSpacingOneofCase.LineSpacingMultiplier ||
        source.SpaceBeforeCase is PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforePoints or PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforeMultiplier ||
        source.SpaceAfterCase is PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterPoints or PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterMultiplier;

    internal static void Append(A.TextParagraphPropertiesType target, PresentationTextParagraph source)
    {
        if (source.LineSpacingCase is PresentationTextParagraph.LineSpacingOneofCase.LineSpacingPoints or PresentationTextParagraph.LineSpacingOneofCase.LineSpacingMultiplier)
            target.AddChild(BuildLineSpacing(source), true);
        if (source.SpaceBeforeCase is PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforePoints or PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforeMultiplier)
            target.AddChild(BuildSpaceBefore(source), true);
        if (source.SpaceAfterCase is PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterPoints or PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterMultiplier)
            target.AddChild(BuildSpaceAfter(source), true);
    }

    internal static void Apply(A.TextParagraphPropertiesType target, PresentationTextParagraph source)
    {
        if (source.LineSpacingCase != PresentationTextParagraph.LineSpacingOneofCase.None)
            Replace(target, target.Elements<A.LineSpacing>(), false,
                source.LineSpacingCase == PresentationTextParagraph.LineSpacingOneofCase.NoLineSpacing ? null : BuildLineSpacing(source), "line spacing");
        if (source.SpaceBeforeCase != PresentationTextParagraph.SpaceBeforeOneofCase.None)
            Replace(target, target.Elements<A.SpaceBefore>(), true,
                source.SpaceBeforeCase == PresentationTextParagraph.SpaceBeforeOneofCase.NoSpaceBefore ? null : BuildSpaceBefore(source), "space before");
        if (source.SpaceAfterCase != PresentationTextParagraph.SpaceAfterOneofCase.None)
            Replace(target, target.Elements<A.SpaceAfter>(), true,
                source.SpaceAfterCase == PresentationTextParagraph.SpaceAfterOneofCase.NoSpaceAfter ? null : BuildSpaceAfter(source), "space after");
    }

    internal static void Scrub(A.TextParagraphPropertiesType target)
    {
        Scrub(target.Elements<A.LineSpacing>(), false);
        Scrub(target.Elements<A.SpaceBefore>(), true);
        Scrub(target.Elements<A.SpaceAfter>(), true);
    }

    private static void ReadSlot(A.TextSpacingType? source, bool allowZero, Action<double> setPoints, Action<double> setMultiplier)
    {
        if (!SupportsSlot(source, allowZero) || source is null) return;
        if (source.GetFirstChild<A.SpacingPoints>()?.Val?.Value is { } points) setPoints(points / 100d);
        else if (source.GetFirstChild<A.SpacingPercent>()?.Val?.Value is { } percent) setMultiplier(percent / 100_000d);
    }

    private static bool SupportsSingle<T>(IEnumerable<T> source, bool allowZero) where T : A.TextSpacingType
    {
        var slots = source.ToArray();
        return slots.Length <= 1 && (slots.Length == 0 || SupportsSlot(slots[0], allowZero));
    }

    private static bool SupportsSlot(A.TextSpacingType? source, bool allowZero)
    {
        if (source is null) return true;
        if (source.ExtendedAttributes.Any() || source.ChildElements.Count != 1) return false;
        return source.FirstChild switch
        {
            A.SpacingPoints points when !points.ExtendedAttributes.Any() && points.Val?.Value is { } value => ValidNative(value, MaxPointsHundredths, allowZero),
            A.SpacingPercent percent when !percent.ExtendedAttributes.Any() && percent.Val?.Value is { } value => ValidNative(value, MaxPercentThousandths, allowZero),
            _ => false,
        };
    }

    private static A.LineSpacing BuildLineSpacing(PresentationTextParagraph source) => new(
        source.LineSpacingCase == PresentationTextParagraph.LineSpacingOneofCase.LineSpacingPoints
            ? new A.SpacingPoints { Val = Points(source.LineSpacingPoints, false, "line spacing") }
            : new A.SpacingPercent { Val = Percent(source.LineSpacingMultiplier, false, "line spacing") });

    private static A.SpaceBefore BuildSpaceBefore(PresentationTextParagraph source) => new(
        source.SpaceBeforeCase == PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforePoints
            ? new A.SpacingPoints { Val = Points(source.SpaceBeforePoints, true, "space before") }
            : new A.SpacingPercent { Val = Percent(source.SpaceBeforeMultiplier, true, "space before") });

    private static A.SpaceAfter BuildSpaceAfter(PresentationTextParagraph source) => new(
        source.SpaceAfterCase == PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterPoints
            ? new A.SpacingPoints { Val = Points(source.SpaceAfterPoints, true, "space after") }
            : new A.SpacingPercent { Val = Percent(source.SpaceAfterMultiplier, true, "space after") });

    private static void Replace<T>(A.TextParagraphPropertiesType target, IEnumerable<T> source, bool allowZero, T? replacement, string kind) where T : A.TextSpacingType
    {
        var slots = source.ToArray();
        if (slots.Length > 1 || slots.Any(slot => !SupportsSlot(slot, allowZero))) throw Unsupported(kind);
        foreach (var slot in slots) slot.Remove();
        if (replacement is not null) target.AddChild(replacement, true);
    }

    private static void Scrub<T>(IEnumerable<T> source, bool allowZero) where T : A.TextSpacingType
    {
        var slots = source.ToArray();
        if (slots.Length == 1 && SupportsSlot(slots[0], allowZero)) slots[0].Remove();
    }

    private static int Points(double value, bool allowZero, string kind) => Native(value, 100, MaxPointsHundredths, allowZero, kind, "points");
    private static int Percent(double value, bool allowZero, string kind) => Native(value, 100_000, MaxPercentThousandths, allowZero, kind, "multiplier");

    private static int Native(double value, int scale, int maximum, bool allowZero, string kind, string unit)
    {
        if (!double.IsFinite(value)) throw Invalid($"Presentation paragraph {kind} {unit} must be finite.");
        var native = Math.Round(value * scale);
        if (native < (allowZero ? 0 : 1) || native > maximum)
            throw Invalid($"Presentation paragraph {kind} {unit} is outside the supported DrawingML range.");
        return checked((int)native);
    }

    private static bool ValidNative(int value, int maximum, bool allowZero) => value >= (allowZero ? 0 : 1) && value <= maximum;
    private static CodecException Invalid(string message) => new("invalid_presentation_text", message);
    private static CodecException Unsupported(string kind) => new("unsupported_presentation_edit", $"Source-preserving PPTX export cannot replace unmodeled paragraph {kind}.");
}
