using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OpenOffice.OpenXmlCodec;

// Owns direct paragraph coordinates whose absence means inheritance. Explicit
// no_* wire choices carry deletion intent across source-preserving edits.
internal static class PptxParagraphLayoutCodec
{
    private const long MaxCoordinateEmu = 51_206_400;

    internal static void Read(PresentationTextParagraph target, A.ParagraphProperties? source)
    {
        if (source?.LeftMargin?.Value is { } margin && ValidMargin(margin)) target.MarginLeftEmu = margin;
        if (source?.Indent?.Value is { } indent && ValidIndent(indent)) target.IndentEmu = indent;
    }

    internal static bool Supports(A.ParagraphProperties? source) =>
        source is null ||
        (source.LeftMargin is null || ValidMargin(source.LeftMargin.Value)) &&
        (source.Indent is null || ValidIndent(source.Indent.Value));

    internal static void Validate(PresentationTextParagraph source)
    {
        switch (source.LeftMarginCase)
        {
            case PresentationTextParagraph.LeftMarginOneofCase.None:
                break;
            case PresentationTextParagraph.LeftMarginOneofCase.MarginLeftEmu:
                if (!ValidMargin(source.MarginLeftEmu)) throw Invalid($"Presentation paragraph left margin must be from 0 through {MaxCoordinateEmu} EMUs.");
                break;
            case PresentationTextParagraph.LeftMarginOneofCase.NoMarginLeft:
                if (!source.NoMarginLeft) throw Invalid("Presentation no_margin_left must be true when selected.");
                break;
            default:
                throw Invalid("Presentation paragraph contains an unknown left-margin case.");
        }

        switch (source.IndentationCase)
        {
            case PresentationTextParagraph.IndentationOneofCase.None:
                break;
            case PresentationTextParagraph.IndentationOneofCase.IndentEmu:
                if (!ValidIndent(source.IndentEmu)) throw Invalid($"Presentation paragraph indent must be from {-MaxCoordinateEmu} through {MaxCoordinateEmu} EMUs.");
                break;
            case PresentationTextParagraph.IndentationOneofCase.NoIndent:
                if (!source.NoIndent) throw Invalid("Presentation no_indent must be true when selected.");
                break;
            default:
                throw Invalid("Presentation paragraph contains an unknown indentation case.");
        }
    }

    internal static bool HasAuthoredLayout(PresentationTextParagraph source) =>
        source.LeftMarginCase == PresentationTextParagraph.LeftMarginOneofCase.MarginLeftEmu ||
        source.IndentationCase == PresentationTextParagraph.IndentationOneofCase.IndentEmu;

    internal static void Append(A.ParagraphProperties target, PresentationTextParagraph source)
    {
        if (source.LeftMarginCase == PresentationTextParagraph.LeftMarginOneofCase.MarginLeftEmu)
            target.LeftMargin = checked((int)source.MarginLeftEmu);
        if (source.IndentationCase == PresentationTextParagraph.IndentationOneofCase.IndentEmu)
            target.Indent = checked((int)source.IndentEmu);
    }

    internal static void Apply(A.ParagraphProperties target, PresentationTextParagraph source)
    {
        if (source.LeftMarginCase != PresentationTextParagraph.LeftMarginOneofCase.None)
        {
            if (target.LeftMargin is not null && !ValidMargin(target.LeftMargin.Value)) throw Unsupported("left margin");
            target.LeftMargin = source.LeftMarginCase == PresentationTextParagraph.LeftMarginOneofCase.MarginLeftEmu
                ? checked((int)source.MarginLeftEmu)
                : null;
        }
        if (source.IndentationCase != PresentationTextParagraph.IndentationOneofCase.None)
        {
            if (target.Indent is not null && !ValidIndent(target.Indent.Value)) throw Unsupported("indent");
            target.Indent = source.IndentationCase == PresentationTextParagraph.IndentationOneofCase.IndentEmu
                ? checked((int)source.IndentEmu)
                : null;
        }
    }

    internal static void Scrub(A.ParagraphProperties target)
    {
        if (target.LeftMargin?.Value is { } margin && ValidMargin(margin)) target.LeftMargin = null;
        if (target.Indent?.Value is { } indent && ValidIndent(indent)) target.Indent = null;
    }

    private static bool ValidMargin(long value) => value is >= 0 and <= MaxCoordinateEmu;
    private static bool ValidIndent(long value) => value is >= -MaxCoordinateEmu and <= MaxCoordinateEmu;
    private static CodecException Invalid(string message) => new("invalid_presentation_text", message);
    private static CodecException Unsupported(string kind) => new("unsupported_presentation_edit", $"Source-preserving PPTX export cannot replace an unmodeled paragraph {kind}.");
}
