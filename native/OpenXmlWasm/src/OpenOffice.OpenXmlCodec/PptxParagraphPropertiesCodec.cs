using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OpenOffice.OpenXmlCodec;

// Coordinates the modeled CT_TextParagraphProperties subset shared by direct
// a:pPr paragraphs and a:lstStyle level defaults. Child-specific codecs retain
// ownership of their native elements and source-preserving edit rules.
internal static class PptxParagraphPropertiesCodec
{
    internal static void Read(
        PresentationTextParagraph target,
        A.TextParagraphPropertiesType? source,
        PptxSlideContext? slideContext,
        bool readLevel)
    {
        if (readLevel && source?.Level is not null) target.Level = checked((uint)source.Level.Value);
        if (source?.Alignment?.Value is { } alignment && AlignmentName(alignment) is { Length: > 0 } name)
            target.Alignment = name;
        PptxParagraphLayoutCodec.Read(target, source);
        PptxParagraphSpacingCodec.Read(target, source);
        PptxBulletCodec.Read(target, source, slideContext);
        PptxBulletStyleCodec.Read(target, source);
        PptxTextCodec.ReadTabStops(target, source);
        PptxDefaultRunStyleCodec.Read(target, source);
    }

    internal static bool Supports(A.TextParagraphPropertiesType? source) =>
        PptxTextCodec.SupportsTabStops(source) &&
        PptxParagraphLayoutCodec.Supports(source) &&
        PptxParagraphSpacingCodec.Supports(source) &&
        PptxDefaultRunStyleCodec.Supports(source);

    internal static void Validate(PresentationTextParagraph source, bool requireLevel)
    {
        if (requireLevel && !source.HasLevel)
            throw Invalid("Presentation list style must identify a level from 0 through 8.");
        if (source.HasLevel && source.Level > 8)
            throw Invalid("Presentation paragraph level must be from 0 through 8.");
        if (source.HasAlignment) _ = ParseAlignment(source.Alignment);
        PptxParagraphLayoutCodec.Validate(source);
        PptxParagraphSpacingCodec.Validate(source);
        PptxDefaultRunStyleCodec.Validate(source);
        PptxBulletCodec.Validate(source);
        PptxBulletStyleCodec.Validate(source);
        PptxTextCodec.ValidateTabStops(source);
    }

    internal static bool HasAuthoredProperties(PresentationTextParagraph source, bool includeLevel) =>
        includeLevel && source.HasLevel ||
        source.HasAlignment ||
        PptxParagraphLayoutCodec.HasAuthoredLayout(source) ||
        PptxParagraphSpacingCodec.HasAuthoredSpacing(source) ||
        PptxBulletCodec.HasModeledBullet(source) ||
        PptxBulletStyleCodec.HasModeledStyle(source) ||
        PptxDefaultRunStyleCodec.HasAuthoredStyle(source) ||
        source.TabStops.Count > 0;

    internal static bool HasModeledProperties(PresentationTextParagraph source) =>
        source.HasAlignment ||
        source.LeftMarginCase != PresentationTextParagraph.LeftMarginOneofCase.None ||
        source.IndentationCase != PresentationTextParagraph.IndentationOneofCase.None ||
        source.LineSpacingCase != PresentationTextParagraph.LineSpacingOneofCase.None ||
        source.SpaceBeforeCase != PresentationTextParagraph.SpaceBeforeOneofCase.None ||
        source.SpaceAfterCase != PresentationTextParagraph.SpaceAfterOneofCase.None ||
        source.BulletCase != PresentationTextParagraph.BulletOneofCase.None ||
        source.BulletFontCase != PresentationTextParagraph.BulletFontOneofCase.None ||
        source.BulletColorCase != PresentationTextParagraph.BulletColorOneofCase.None ||
        source.BulletSizeCase != PresentationTextParagraph.BulletSizeOneofCase.None ||
        source.TabStops.Count > 0 || source.HasNoTabStops ||
        source.DefaultRunStyleCase != PresentationTextParagraph.DefaultRunStyleOneofCase.None;

    internal static void Append(
        A.TextParagraphPropertiesType target,
        PresentationTextParagraph source,
        PptxSlideContext? slideContext,
        bool includeLevel)
    {
        if (includeLevel && source.HasLevel) target.Level = checked((int)source.Level);
        if (source.HasAlignment) target.Alignment = ParseAlignment(source.Alignment);
        PptxParagraphLayoutCodec.Append(target, source);
        PptxParagraphSpacingCodec.Append(target, source);
        PptxBulletStyleCodec.Append(target, source);
        PptxBulletCodec.Append(target, source, slideContext);
        PptxTextCodec.AppendTabStops(target, source);
        PptxDefaultRunStyleCodec.Append(target, source);
    }

    internal static void Apply(
        A.TextParagraphPropertiesType target,
        PresentationTextParagraph source,
        PptxSlideContext slideContext,
        bool includeLevel)
    {
        if (includeLevel) target.Level = source.HasLevel ? checked((int)source.Level) : null;
        if (source.HasAlignment) target.Alignment = ParseAlignment(source.Alignment);
        else if (target.Alignment?.Value is { } alignment && AlignmentName(alignment).Length > 0) target.Alignment = null;
        PptxParagraphLayoutCodec.Apply(target, source);
        PptxParagraphSpacingCodec.Apply(target, source);
        PptxBulletStyleCodec.Apply(target, source);
        PptxBulletCodec.Apply(target, source, slideContext);
        PptxTextCodec.ApplyTabStops(target, source);
        PptxDefaultRunStyleCodec.Apply(target, source);
    }

    internal static void Scrub(A.TextParagraphPropertiesType target, PptxSlideContext? slideContext, bool includeLevel)
    {
        if (includeLevel) target.Level = null;
        if (target.Alignment?.Value is { } alignment && AlignmentName(alignment).Length > 0) target.Alignment = null;
        PptxParagraphLayoutCodec.Scrub(target);
        PptxParagraphSpacingCodec.Scrub(target);
        PptxDefaultRunStyleCodec.Scrub(target);
        PptxBulletCodec.Scrub(target, slideContext);
        PptxBulletStyleCodec.Scrub(target);
        target.GetFirstChild<A.TabStopList>()?.Remove();
    }

    private static string AlignmentName(A.TextAlignmentTypeValues value) =>
        value == A.TextAlignmentTypeValues.Left ? "left" :
        value == A.TextAlignmentTypeValues.Center ? "center" :
        value == A.TextAlignmentTypeValues.Right ? "right" :
        value == A.TextAlignmentTypeValues.Justified ? "justify" : string.Empty;

    private static A.TextAlignmentTypeValues ParseAlignment(string value) => value switch
    {
        "left" => A.TextAlignmentTypeValues.Left,
        "center" => A.TextAlignmentTypeValues.Center,
        "right" => A.TextAlignmentTypeValues.Right,
        "justify" => A.TextAlignmentTypeValues.Justified,
        _ => throw Invalid($"Unsupported Presentation paragraph alignment {value}."),
    };

    private static CodecException Invalid(string message) => new("invalid_presentation_text", message);
}
