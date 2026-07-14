using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenChestnut.Codec;

// Owns only a direct, bounded p:bg choice on one p:cSld. Effective color
// inheritance stays in the JavaScript model; gradients, patterns, images,
// transforms, and effect-bearing backgrounds remain source-bound.
internal static class PptxBackgroundCodec
{
    internal static PresentationBackground? Read(P.CommonSlideData? source)
    {
        var background = source?.GetFirstChild<P.Background>();
        return background is not null && TryRead(background, out var semantic) ? semantic : null;
    }

    internal static bool Supports(P.CommonSlideData? source)
    {
        var backgrounds = source?.Elements<P.Background>().ToArray() ?? [];
        return backgrounds.Length == 0 || backgrounds.Length == 1 && TryRead(backgrounds[0], out _);
    }

    internal static void Validate(PresentationBackground? source)
    {
        if (source is null) return;
        switch (source.ColorCase)
        {
            case PresentationBackground.ColorOneofCase.ColorRgb:
                _ = PptxColor.Normalize(source.ColorRgb);
                break;
            case PresentationBackground.ColorOneofCase.ColorScheme:
                _ = PptxColor.NormalizeScheme(source.ColorScheme);
                break;
            default:
                throw Invalid("Presentation background requires exactly one RGB or theme color.");
        }
        switch (source.KindCase)
        {
            case PresentationBackground.KindOneofCase.Solid when source.Solid:
                break;
            case PresentationBackground.KindOneofCase.StyleReferenceIndex:
                break;
            default:
                throw Invalid("Presentation background requires a solid mode or style-reference index.");
        }
    }

    internal static void Build(P.CommonSlideData target, PresentationBackground? source)
    {
        if (source is null) return;
        target.AddChild(BuildElement(source), true);
    }

    internal static void Apply(P.CommonSlideData target, PresentationBackground source)
    {
        Validate(source);
        var current = target.GetFirstChild<P.Background>();
        var replacement = BuildElement(source);
        if (current is null)
        {
            target.AddChild(replacement, true);
            return;
        }
        current.InsertAfterSelf(replacement);
        current.Remove();
    }

    internal static void ScrubModeledContent(P.CommonSlideData? source)
    {
        var background = source?.GetFirstChild<P.Background>();
        if (background is not null && TryRead(background, out _)) background.Remove();
    }

    private static bool TryRead(P.Background source, out PresentationBackground semantic)
    {
        semantic = new PresentationBackground();
        if (source.GetAttributes().Count != 0 || source.ChildElements.Count != 1) return false;
        switch (source.FirstChild)
        {
            case P.BackgroundProperties properties:
                if (properties.GetAttributes().Count != 0) return false;
                var children = properties.ChildElements.ToArray();
                if (children.Length is < 1 or > 2 || children[0] is not A.SolidFill solid) return false;
                if (children.Length == 2 && (children[1] is not A.EffectList { ChildElements.Count: 0 } effectList || effectList.GetAttributes().Count != 0)) return false;
                if (!TryReadColor(solid, out semantic)) return false;
                semantic.Solid = true;
                return true;
            case P.BackgroundStyleReference reference:
                if (reference.Index?.Value is not { } index ||
                    reference.GetAttributes().Any(attribute => attribute.LocalName != "idx") ||
                    !TryReadColor(reference, out semantic)) return false;
                semantic.StyleReferenceIndex = index;
                return true;
            default:
                return false;
        }
    }

    private static bool TryReadColor(OpenXmlCompositeElement source, out PresentationBackground target)
    {
        target = new PresentationBackground();
        if ((source is A.SolidFill && source.GetAttributes().Count != 0) || source.ChildElements.Count != 1) return false;
        switch (source.FirstChild)
        {
            case A.RgbColorModelHex rgb when rgb.ChildElements.Count == 0 && rgb.GetAttributes().All(attribute => attribute.LocalName == "val") && rgb.Val?.Value is { Length: 6 } value:
                try { target.ColorRgb = PptxColor.Normalize(value); return true; }
                catch (CodecException) { return false; }
            case A.SchemeColor scheme when scheme.ChildElements.Count == 0 && scheme.GetAttributes().All(attribute => attribute.LocalName == "val") && scheme.Val?.Value is { } value && PptxColor.TrySchemeToken(value, out var token):
                target.ColorScheme = token;
                return true;
            default:
                return false;
        }
    }

    private static P.Background BuildElement(PresentationBackground source)
    {
        Validate(source);
        return source.KindCase switch
        {
            PresentationBackground.KindOneofCase.Solid =>
                new P.Background(new P.BackgroundProperties(new A.SolidFill(Color(source)), new A.EffectList())),
            PresentationBackground.KindOneofCase.StyleReferenceIndex =>
                new P.Background(new P.BackgroundStyleReference(Color(source)) { Index = source.StyleReferenceIndex }),
            _ => throw Invalid("Presentation background kind is missing."),
        };
    }

    private static OpenXmlElement Color(PresentationBackground source) => source.ColorCase switch
    {
        PresentationBackground.ColorOneofCase.ColorRgb => new A.RgbColorModelHex { Val = PptxColor.Normalize(source.ColorRgb) },
        PresentationBackground.ColorOneofCase.ColorScheme => new A.SchemeColor { Val = PptxColor.SchemeValue(source.ColorScheme) },
        _ => throw Invalid("Presentation background color is missing."),
    };

    private static CodecException Invalid(string message) => new("invalid_presentation_background", message);
}
