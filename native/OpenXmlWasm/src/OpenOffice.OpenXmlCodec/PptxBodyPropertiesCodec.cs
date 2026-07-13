using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenOffice.OpenXmlCodec;

// Owns the bounded a:bodyPr layout subset. Unmodeled attributes and children
// remain on the source element. Attributed AutoFit choices stay source-bound
// because their scaling parameters are not yet part of the public wire model.
internal static class PptxBodyPropertiesCodec
{
    internal static void Read(PresentationTextBody target, P.TextBody source)
    {
        var native = source.Elements<A.BodyProperties>().FirstOrDefault();
        if (native is null) return;
        var modeled = new PresentationTextBodyProperties();
        ReadInset(native.LeftInset?.Value, value => modeled.LeftInsetEmu = value);
        ReadInset(native.TopInset?.Value, value => modeled.TopInsetEmu = value);
        ReadInset(native.RightInset?.Value, value => modeled.RightInsetEmu = value);
        ReadInset(native.BottomInset?.Value, value => modeled.BottomInsetEmu = value);
        if (AnchorName(native.Anchor?.Value) is { Length: > 0 } anchor) modeled.VerticalAnchor = anchor;
        if (WrapName(native.Wrap?.Value) is { Length: > 0 } wrap) modeled.Wrap = wrap;
        var autoFit = native.ChildElements.Where(IsAutoFitChoice).ToArray();
        if (autoFit.Length == 1 && IsSimple(autoFit[0]) && AutoFitName(autoFit[0]) is { Length: > 0 } mode) modeled.AutoFitMode = mode;
        if (HasModeledProperties(modeled)) target.BodyProperties = modeled;
    }

    internal static bool Supports(P.TextBody? source)
    {
        if (source is null) return true;
        var bodies = source.Elements<A.BodyProperties>().ToArray();
        return bodies.Length <= 1 && (bodies.Length == 0 || bodies[0].ChildElements.Count(IsAutoFitChoice) <= 1);
    }

    internal static void Validate(PresentationTextBody source)
    {
        if (source.BodyProperties is null) return;
        var properties = source.BodyProperties;
        ValidateInset(properties.LeftInsetCase, properties.LeftInsetEmu, PresentationTextBodyProperties.LeftInsetOneofCase.LeftInsetEmu, PresentationTextBodyProperties.LeftInsetOneofCase.NoLeftInset, properties.NoLeftInset, "left");
        ValidateInset(properties.TopInsetCase, properties.TopInsetEmu, PresentationTextBodyProperties.TopInsetOneofCase.TopInsetEmu, PresentationTextBodyProperties.TopInsetOneofCase.NoTopInset, properties.NoTopInset, "top");
        ValidateInset(properties.RightInsetCase, properties.RightInsetEmu, PresentationTextBodyProperties.RightInsetOneofCase.RightInsetEmu, PresentationTextBodyProperties.RightInsetOneofCase.NoRightInset, properties.NoRightInset, "right");
        ValidateInset(properties.BottomInsetCase, properties.BottomInsetEmu, PresentationTextBodyProperties.BottomInsetOneofCase.BottomInsetEmu, PresentationTextBodyProperties.BottomInsetOneofCase.NoBottomInset, properties.NoBottomInset, "bottom");
        if (properties.AnchorCase == PresentationTextBodyProperties.AnchorOneofCase.VerticalAnchor) _ = ParseAnchor(properties.VerticalAnchor);
        else if (properties.AnchorCase == PresentationTextBodyProperties.AnchorOneofCase.NoVerticalAnchor && !properties.NoVerticalAnchor) throw Invalid("Presentation no_vertical_anchor must be true when selected.");
        if (properties.WrappingCase == PresentationTextBodyProperties.WrappingOneofCase.Wrap) _ = ParseWrap(properties.Wrap);
        else if (properties.WrappingCase == PresentationTextBodyProperties.WrappingOneofCase.NoWrap && !properties.NoWrap) throw Invalid("Presentation no_wrap must be true when selected.");
        if (properties.AutoFitCase == PresentationTextBodyProperties.AutoFitOneofCase.AutoFitMode) _ = ParseAutoFit(properties.AutoFitMode);
        else if (properties.AutoFitCase == PresentationTextBodyProperties.AutoFitOneofCase.NoAutoFitMode && !properties.NoAutoFitMode) throw Invalid("Presentation no_auto_fit_mode must be true when selected.");
    }

    internal static bool HasModeledProperties(PresentationTextBodyProperties? source) => source is not null &&
        (source.LeftInsetCase != PresentationTextBodyProperties.LeftInsetOneofCase.None ||
         source.TopInsetCase != PresentationTextBodyProperties.TopInsetOneofCase.None ||
         source.RightInsetCase != PresentationTextBodyProperties.RightInsetOneofCase.None ||
         source.BottomInsetCase != PresentationTextBodyProperties.BottomInsetOneofCase.None ||
         source.AnchorCase != PresentationTextBodyProperties.AnchorOneofCase.None ||
         source.WrappingCase != PresentationTextBodyProperties.WrappingOneofCase.None ||
         source.AutoFitCase != PresentationTextBodyProperties.AutoFitOneofCase.None);

    internal static void Build(A.BodyProperties target, PresentationTextBody source)
    {
        if (source.BodyProperties is not { } properties) return;
        ApplyInsets(target, properties);
        if (properties.AnchorCase == PresentationTextBodyProperties.AnchorOneofCase.VerticalAnchor) target.Anchor = ParseAnchor(properties.VerticalAnchor);
        if (properties.WrappingCase == PresentationTextBodyProperties.WrappingOneofCase.Wrap) target.Wrap = ParseWrap(properties.Wrap);
        if (properties.AutoFitCase == PresentationTextBodyProperties.AutoFitOneofCase.AutoFitMode) target.AddChild(CreateAutoFit(properties.AutoFitMode), true);
    }

    internal static void Apply(P.TextBody target, PresentationTextBody source)
    {
        var bodies = target.Elements<A.BodyProperties>().ToArray();
        if (bodies.Length > 1) throw Unsupported("Source-preserving PPTX export cannot edit duplicate text body properties.");
        var native = bodies.FirstOrDefault();
        if (native is null)
        {
            if (!HasModeledProperties(source.BodyProperties)) return;
            native = new A.BodyProperties();
            target.PrependChild(native);
        }
        var properties = source.BodyProperties;
        if (properties is null) return;
        ApplyInset(properties.LeftInsetCase == PresentationTextBodyProperties.LeftInsetOneofCase.LeftInsetEmu, properties.LeftInsetCase == PresentationTextBodyProperties.LeftInsetOneofCase.NoLeftInset, properties.LeftInsetEmu, value => native.LeftInset = value, () => native.LeftInset = null);
        ApplyInset(properties.TopInsetCase == PresentationTextBodyProperties.TopInsetOneofCase.TopInsetEmu, properties.TopInsetCase == PresentationTextBodyProperties.TopInsetOneofCase.NoTopInset, properties.TopInsetEmu, value => native.TopInset = value, () => native.TopInset = null);
        ApplyInset(properties.RightInsetCase == PresentationTextBodyProperties.RightInsetOneofCase.RightInsetEmu, properties.RightInsetCase == PresentationTextBodyProperties.RightInsetOneofCase.NoRightInset, properties.RightInsetEmu, value => native.RightInset = value, () => native.RightInset = null);
        ApplyInset(properties.BottomInsetCase == PresentationTextBodyProperties.BottomInsetOneofCase.BottomInsetEmu, properties.BottomInsetCase == PresentationTextBodyProperties.BottomInsetOneofCase.NoBottomInset, properties.BottomInsetEmu, value => native.BottomInset = value, () => native.BottomInset = null);
        if (properties.AnchorCase == PresentationTextBodyProperties.AnchorOneofCase.VerticalAnchor) native.Anchor = ParseAnchor(properties.VerticalAnchor);
        else if (properties.AnchorCase == PresentationTextBodyProperties.AnchorOneofCase.NoVerticalAnchor) native.Anchor = null;
        if (properties.WrappingCase == PresentationTextBodyProperties.WrappingOneofCase.Wrap) native.Wrap = ParseWrap(properties.Wrap);
        else if (properties.WrappingCase == PresentationTextBodyProperties.WrappingOneofCase.NoWrap) native.Wrap = null;
        ApplyAutoFit(native, properties);
    }

    internal static void Scrub(P.TextBody? source)
    {
        foreach (var native in source?.Elements<A.BodyProperties>() ?? [])
        {
            if (native.LeftInset?.Value is >= 0) native.LeftInset = null;
            if (native.TopInset?.Value is >= 0) native.TopInset = null;
            if (native.RightInset?.Value is >= 0) native.RightInset = null;
            if (native.BottomInset?.Value is >= 0) native.BottomInset = null;
            if (AnchorName(native.Anchor?.Value).Length > 0) native.Anchor = null;
            if (WrapName(native.Wrap?.Value).Length > 0) native.Wrap = null;
            foreach (var autoFit in native.ChildElements.Where(child => IsAutoFitChoice(child) && IsSimple(child)).ToArray()) autoFit.Remove();
        }
    }

    private static void ReadInset(int? value, Action<long> assign)
    {
        if (value is >= 0) assign(value.Value);
    }

    private static void ValidateInset<TCase>(TCase actualCase, long value, TCase valueCase, TCase noCase, bool noValue, string name) where TCase : struct, Enum
    {
        if (EqualityComparer<TCase>.Default.Equals(actualCase, valueCase) && (value < 0 || value > int.MaxValue)) throw Invalid($"Presentation {name} text inset must fit the non-negative signed 32-bit EMU range.");
        if (EqualityComparer<TCase>.Default.Equals(actualCase, noCase) && !noValue) throw Invalid($"Presentation no_{name}_inset must be true when selected.");
    }

    private static void ApplyInsets(A.BodyProperties target, PresentationTextBodyProperties source)
    {
        if (source.LeftInsetCase == PresentationTextBodyProperties.LeftInsetOneofCase.LeftInsetEmu) target.LeftInset = checked((int)source.LeftInsetEmu);
        if (source.TopInsetCase == PresentationTextBodyProperties.TopInsetOneofCase.TopInsetEmu) target.TopInset = checked((int)source.TopInsetEmu);
        if (source.RightInsetCase == PresentationTextBodyProperties.RightInsetOneofCase.RightInsetEmu) target.RightInset = checked((int)source.RightInsetEmu);
        if (source.BottomInsetCase == PresentationTextBodyProperties.BottomInsetOneofCase.BottomInsetEmu) target.BottomInset = checked((int)source.BottomInsetEmu);
    }

    private static void ApplyInset(bool hasValue, bool hasDelete, long value, Action<int> set, Action clear)
    {
        if (hasValue) set(checked((int)value));
        else if (hasDelete) clear();
    }

    private static void ApplyAutoFit(A.BodyProperties target, PresentationTextBodyProperties source)
    {
        if (source.AutoFitCase == PresentationTextBodyProperties.AutoFitOneofCase.None) return;
        var choices = target.ChildElements.Where(IsAutoFitChoice).ToArray();
        if (choices.Length > 1) throw Unsupported("Source-preserving PPTX export cannot replace duplicate AutoFit choices.");
        var current = choices.FirstOrDefault();
        if (current is not null && !IsSimple(current)) throw Unsupported("Source-preserving PPTX export cannot replace an attributed AutoFit choice.");
        if (source.AutoFitCase == PresentationTextBodyProperties.AutoFitOneofCase.NoAutoFitMode)
        {
            current?.Remove();
            return;
        }
        var mode = source.AutoFitMode;
        if (current is not null && AutoFitName(current) == mode) return;
        current?.Remove();
        target.AddChild(CreateAutoFit(mode), true);
    }

    private static bool IsAutoFitChoice(OpenXmlElement child) => child is A.NoAutoFit or A.NormalAutoFit or A.ShapeAutoFit;
    private static bool IsSimple(OpenXmlElement child) => child.GetAttributes().Count == 0 && child.ChildElements.Count == 0;
    private static string AutoFitName(OpenXmlElement child) => child switch
    {
        A.NoAutoFit => "none",
        A.NormalAutoFit => "shrinkText",
        A.ShapeAutoFit => "resizeShape",
        _ => string.Empty,
    };

    private static OpenXmlElement CreateAutoFit(string value) => ParseAutoFit(value) switch
    {
        "none" => new A.NoAutoFit(),
        "shrinkText" => new A.NormalAutoFit(),
        "resizeShape" => new A.ShapeAutoFit(),
        _ => throw Invalid($"Unsupported Presentation AutoFit mode {value}."),
    };

    private static string ParseAutoFit(string value) => value switch
    {
        "none" or "shrinkText" or "resizeShape" => value,
        _ => throw Invalid($"Unsupported Presentation AutoFit mode {value}."),
    };

    private static string AnchorName(A.TextAnchoringTypeValues? value) => value is null ? string.Empty :
        value.Value == A.TextAnchoringTypeValues.Top ? "top" :
        value.Value == A.TextAnchoringTypeValues.Center ? "center" :
        value.Value == A.TextAnchoringTypeValues.Bottom ? "bottom" : string.Empty;

    private static A.TextAnchoringTypeValues ParseAnchor(string value) => value switch
    {
        "top" => A.TextAnchoringTypeValues.Top,
        "center" => A.TextAnchoringTypeValues.Center,
        "bottom" => A.TextAnchoringTypeValues.Bottom,
        _ => throw Invalid($"Unsupported Presentation text body anchor {value}."),
    };

    private static string WrapName(A.TextWrappingValues? value) => value is null ? string.Empty :
        value.Value == A.TextWrappingValues.Square ? "square" :
        value.Value == A.TextWrappingValues.None ? "none" : string.Empty;

    private static A.TextWrappingValues ParseWrap(string value) => value switch
    {
        "square" => A.TextWrappingValues.Square,
        "none" => A.TextWrappingValues.None,
        _ => throw Invalid($"Unsupported Presentation text body wrap mode {value}."),
    };

    private static CodecException Invalid(string message) => new("invalid_presentation_text", message);
    private static CodecException Unsupported(string message) => new("unsupported_presentation_edit", message);
}
