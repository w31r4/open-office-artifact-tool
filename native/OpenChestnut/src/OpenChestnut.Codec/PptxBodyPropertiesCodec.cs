using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenChestnut.Codec;

// Owns the bounded a:bodyPr layout subset. Unmodeled attributes and children
// remain on the source element. Attributed AutoFit choices stay source-bound
// because their scaling parameters are not yet part of the public wire model.
internal static class PptxBodyPropertiesCodec
{
    private const int MaxRotationAngle60000 = 21_600_000;

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
        if (native.Rotation?.Value is >= -MaxRotationAngle60000 and <= MaxRotationAngle60000) modeled.RotationAngle60000 = native.Rotation.Value;
        if (VerticalTextName(native.Vertical?.Value) is { Length: > 0 } verticalText) modeled.VerticalTextMode = verticalText;
        if (VerticalOverflowName(native.VerticalOverflow?.Value) is { Length: > 0 } verticalOverflow) modeled.VerticalOverflowMode = verticalOverflow;
        if (HorizontalOverflowName(native.HorizontalOverflow?.Value) is { Length: > 0 } horizontalOverflow) modeled.HorizontalOverflowMode = horizontalOverflow;
        if (native.ColumnCount?.Value is >= 1 and <= 16) modeled.Columns = checked((uint)native.ColumnCount.Value);
        if (native.ColumnSpacing?.Value is >= 0) modeled.ColumnSpacingEmu = native.ColumnSpacing.Value;
        if (native.RightToLeftColumns?.Value is { } rightToLeft) modeled.RightToLeftColumns = rightToLeft;
        if (native.UpRight?.Value is { } upright) modeled.Upright = upright;
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
        if (properties.RotationCase == PresentationTextBodyProperties.RotationOneofCase.RotationAngle60000 && Math.Abs((long)properties.RotationAngle60000) > MaxRotationAngle60000) throw Invalid("Presentation text body rotation must be between -360 and 360 degrees.");
        else if (properties.RotationCase == PresentationTextBodyProperties.RotationOneofCase.NoRotation && !properties.NoRotation) throw Invalid("Presentation no_rotation must be true when selected.");
        if (properties.VerticalTextCase == PresentationTextBodyProperties.VerticalTextOneofCase.VerticalTextMode) _ = ParseVerticalText(properties.VerticalTextMode);
        else if (properties.VerticalTextCase == PresentationTextBodyProperties.VerticalTextOneofCase.NoVerticalTextMode && !properties.NoVerticalTextMode) throw Invalid("Presentation no_vertical_text_mode must be true when selected.");
        if (properties.VerticalOverflowCase == PresentationTextBodyProperties.VerticalOverflowOneofCase.VerticalOverflowMode) _ = ParseVerticalOverflow(properties.VerticalOverflowMode);
        else if (properties.VerticalOverflowCase == PresentationTextBodyProperties.VerticalOverflowOneofCase.NoVerticalOverflowMode && !properties.NoVerticalOverflowMode) throw Invalid("Presentation no_vertical_overflow_mode must be true when selected.");
        if (properties.HorizontalOverflowCase == PresentationTextBodyProperties.HorizontalOverflowOneofCase.HorizontalOverflowMode) _ = ParseHorizontalOverflow(properties.HorizontalOverflowMode);
        else if (properties.HorizontalOverflowCase == PresentationTextBodyProperties.HorizontalOverflowOneofCase.NoHorizontalOverflowMode && !properties.NoHorizontalOverflowMode) throw Invalid("Presentation no_horizontal_overflow_mode must be true when selected.");
        if (properties.ColumnCountCase == PresentationTextBodyProperties.ColumnCountOneofCase.Columns && (properties.Columns < 1 || properties.Columns > 16)) throw Invalid("Presentation text body column count must be from 1 through 16.");
        else if (properties.ColumnCountCase == PresentationTextBodyProperties.ColumnCountOneofCase.NoColumns && !properties.NoColumns) throw Invalid("Presentation no_columns must be true when selected.");
        if (properties.ColumnSpacingCase == PresentationTextBodyProperties.ColumnSpacingOneofCase.ColumnSpacingEmu && (properties.ColumnSpacingEmu < 0 || properties.ColumnSpacingEmu > int.MaxValue)) throw Invalid("Presentation text body column spacing must fit the non-negative signed 32-bit EMU range.");
        else if (properties.ColumnSpacingCase == PresentationTextBodyProperties.ColumnSpacingOneofCase.NoColumnSpacing && !properties.NoColumnSpacing) throw Invalid("Presentation no_column_spacing must be true when selected.");
        if (properties.ColumnDirectionCase == PresentationTextBodyProperties.ColumnDirectionOneofCase.NoColumnDirection && !properties.NoColumnDirection) throw Invalid("Presentation no_column_direction must be true when selected.");
        if (properties.UprightTextCase == PresentationTextBodyProperties.UprightTextOneofCase.NoUpright && !properties.NoUpright) throw Invalid("Presentation no_upright must be true when selected.");
    }

    internal static bool HasModeledProperties(PresentationTextBodyProperties? source) => source is not null &&
        (source.LeftInsetCase != PresentationTextBodyProperties.LeftInsetOneofCase.None ||
         source.TopInsetCase != PresentationTextBodyProperties.TopInsetOneofCase.None ||
         source.RightInsetCase != PresentationTextBodyProperties.RightInsetOneofCase.None ||
         source.BottomInsetCase != PresentationTextBodyProperties.BottomInsetOneofCase.None ||
         source.AnchorCase != PresentationTextBodyProperties.AnchorOneofCase.None ||
         source.WrappingCase != PresentationTextBodyProperties.WrappingOneofCase.None ||
         source.AutoFitCase != PresentationTextBodyProperties.AutoFitOneofCase.None ||
         source.RotationCase != PresentationTextBodyProperties.RotationOneofCase.None ||
         source.VerticalTextCase != PresentationTextBodyProperties.VerticalTextOneofCase.None ||
         source.VerticalOverflowCase != PresentationTextBodyProperties.VerticalOverflowOneofCase.None ||
         source.HorizontalOverflowCase != PresentationTextBodyProperties.HorizontalOverflowOneofCase.None ||
         source.ColumnCountCase != PresentationTextBodyProperties.ColumnCountOneofCase.None ||
         source.ColumnSpacingCase != PresentationTextBodyProperties.ColumnSpacingOneofCase.None ||
         source.ColumnDirectionCase != PresentationTextBodyProperties.ColumnDirectionOneofCase.None ||
         source.UprightTextCase != PresentationTextBodyProperties.UprightTextOneofCase.None);

    internal static void Build(A.BodyProperties target, PresentationTextBody source)
    {
        if (source.BodyProperties is not { } properties) return;
        ApplyInsets(target, properties);
        if (properties.AnchorCase == PresentationTextBodyProperties.AnchorOneofCase.VerticalAnchor) target.Anchor = ParseAnchor(properties.VerticalAnchor);
        if (properties.WrappingCase == PresentationTextBodyProperties.WrappingOneofCase.Wrap) target.Wrap = ParseWrap(properties.Wrap);
        if (properties.RotationCase == PresentationTextBodyProperties.RotationOneofCase.RotationAngle60000) target.Rotation = properties.RotationAngle60000;
        if (properties.VerticalTextCase == PresentationTextBodyProperties.VerticalTextOneofCase.VerticalTextMode) target.Vertical = ParseVerticalText(properties.VerticalTextMode);
        if (properties.VerticalOverflowCase == PresentationTextBodyProperties.VerticalOverflowOneofCase.VerticalOverflowMode) target.VerticalOverflow = ParseVerticalOverflow(properties.VerticalOverflowMode);
        if (properties.HorizontalOverflowCase == PresentationTextBodyProperties.HorizontalOverflowOneofCase.HorizontalOverflowMode) target.HorizontalOverflow = ParseHorizontalOverflow(properties.HorizontalOverflowMode);
        if (properties.ColumnCountCase == PresentationTextBodyProperties.ColumnCountOneofCase.Columns) target.ColumnCount = checked((int)properties.Columns);
        if (properties.ColumnSpacingCase == PresentationTextBodyProperties.ColumnSpacingOneofCase.ColumnSpacingEmu) target.ColumnSpacing = checked((int)properties.ColumnSpacingEmu);
        if (properties.ColumnDirectionCase == PresentationTextBodyProperties.ColumnDirectionOneofCase.RightToLeftColumns) target.RightToLeftColumns = properties.RightToLeftColumns;
        if (properties.UprightTextCase == PresentationTextBodyProperties.UprightTextOneofCase.Upright) target.UpRight = properties.Upright;
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
        if (properties.RotationCase == PresentationTextBodyProperties.RotationOneofCase.RotationAngle60000) native.Rotation = properties.RotationAngle60000;
        else if (properties.RotationCase == PresentationTextBodyProperties.RotationOneofCase.NoRotation) native.Rotation = null;
        if (properties.VerticalTextCase == PresentationTextBodyProperties.VerticalTextOneofCase.VerticalTextMode) native.Vertical = ParseVerticalText(properties.VerticalTextMode);
        else if (properties.VerticalTextCase == PresentationTextBodyProperties.VerticalTextOneofCase.NoVerticalTextMode) native.Vertical = null;
        if (properties.VerticalOverflowCase == PresentationTextBodyProperties.VerticalOverflowOneofCase.VerticalOverflowMode) native.VerticalOverflow = ParseVerticalOverflow(properties.VerticalOverflowMode);
        else if (properties.VerticalOverflowCase == PresentationTextBodyProperties.VerticalOverflowOneofCase.NoVerticalOverflowMode) native.VerticalOverflow = null;
        if (properties.HorizontalOverflowCase == PresentationTextBodyProperties.HorizontalOverflowOneofCase.HorizontalOverflowMode) native.HorizontalOverflow = ParseHorizontalOverflow(properties.HorizontalOverflowMode);
        else if (properties.HorizontalOverflowCase == PresentationTextBodyProperties.HorizontalOverflowOneofCase.NoHorizontalOverflowMode) native.HorizontalOverflow = null;
        if (properties.ColumnCountCase == PresentationTextBodyProperties.ColumnCountOneofCase.Columns) native.ColumnCount = checked((int)properties.Columns);
        else if (properties.ColumnCountCase == PresentationTextBodyProperties.ColumnCountOneofCase.NoColumns) native.ColumnCount = null;
        if (properties.ColumnSpacingCase == PresentationTextBodyProperties.ColumnSpacingOneofCase.ColumnSpacingEmu) native.ColumnSpacing = checked((int)properties.ColumnSpacingEmu);
        else if (properties.ColumnSpacingCase == PresentationTextBodyProperties.ColumnSpacingOneofCase.NoColumnSpacing) native.ColumnSpacing = null;
        if (properties.ColumnDirectionCase == PresentationTextBodyProperties.ColumnDirectionOneofCase.RightToLeftColumns) native.RightToLeftColumns = properties.RightToLeftColumns;
        else if (properties.ColumnDirectionCase == PresentationTextBodyProperties.ColumnDirectionOneofCase.NoColumnDirection) native.RightToLeftColumns = null;
        if (properties.UprightTextCase == PresentationTextBodyProperties.UprightTextOneofCase.Upright) native.UpRight = properties.Upright;
        else if (properties.UprightTextCase == PresentationTextBodyProperties.UprightTextOneofCase.NoUpright) native.UpRight = null;
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
            if (native.Rotation?.Value is >= -MaxRotationAngle60000 and <= MaxRotationAngle60000) native.Rotation = null;
            if (VerticalTextName(native.Vertical?.Value).Length > 0) native.Vertical = null;
            if (VerticalOverflowName(native.VerticalOverflow?.Value).Length > 0) native.VerticalOverflow = null;
            if (HorizontalOverflowName(native.HorizontalOverflow?.Value).Length > 0) native.HorizontalOverflow = null;
            if (native.ColumnCount?.Value is >= 1 and <= 16) native.ColumnCount = null;
            if (native.ColumnSpacing?.Value is >= 0) native.ColumnSpacing = null;
            if (native.RightToLeftColumns is not null) native.RightToLeftColumns = null;
            if (native.UpRight is not null) native.UpRight = null;
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

    private static string VerticalTextName(A.TextVerticalValues? value) => value is null ? string.Empty :
        value.Value == A.TextVerticalValues.Horizontal ? "horizontal" :
        value.Value == A.TextVerticalValues.Vertical ? "vertical" :
        value.Value == A.TextVerticalValues.Vertical270 ? "vertical270" : string.Empty;

    private static A.TextVerticalValues ParseVerticalText(string value) => value switch
    {
        "horizontal" => A.TextVerticalValues.Horizontal,
        "vertical" => A.TextVerticalValues.Vertical,
        "vertical270" => A.TextVerticalValues.Vertical270,
        _ => throw Invalid($"Unsupported Presentation vertical text mode {value}."),
    };

    private static string VerticalOverflowName(A.TextVerticalOverflowValues? value) => value is null ? string.Empty :
        value.Value == A.TextVerticalOverflowValues.Overflow ? "overflow" :
        value.Value == A.TextVerticalOverflowValues.Ellipsis ? "ellipsis" :
        value.Value == A.TextVerticalOverflowValues.Clip ? "clip" : string.Empty;

    private static A.TextVerticalOverflowValues ParseVerticalOverflow(string value) => value switch
    {
        "overflow" => A.TextVerticalOverflowValues.Overflow,
        "ellipsis" => A.TextVerticalOverflowValues.Ellipsis,
        "clip" => A.TextVerticalOverflowValues.Clip,
        _ => throw Invalid($"Unsupported Presentation vertical overflow mode {value}."),
    };

    private static string HorizontalOverflowName(A.TextHorizontalOverflowValues? value) => value is null ? string.Empty :
        value.Value == A.TextHorizontalOverflowValues.Overflow ? "overflow" :
        value.Value == A.TextHorizontalOverflowValues.Clip ? "clip" : string.Empty;

    private static A.TextHorizontalOverflowValues ParseHorizontalOverflow(string value) => value switch
    {
        "overflow" => A.TextHorizontalOverflowValues.Overflow,
        "clip" => A.TextHorizontalOverflowValues.Clip,
        _ => throw Invalid($"Unsupported Presentation horizontal overflow mode {value}."),
    };

    private static CodecException Invalid(string message) => new("invalid_presentation_text", message);
    private static CodecException Unsupported(string message) => new("unsupported_presentation_edit", message);
}
