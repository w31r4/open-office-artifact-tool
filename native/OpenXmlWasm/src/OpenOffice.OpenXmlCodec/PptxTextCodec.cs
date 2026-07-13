using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenOffice.OpenXmlCodec;

// Owns the public PresentationML paragraph/run subset. Package ownership,
// slide bindings, and opaque graph preservation remain in PptxCodec.
internal static class PptxTextCodec
{
    private const int MaxParagraphs = 4_096;
    private const int MaxRuns = 16_384;
    private const double MaxFontSizePoints = 768;

    internal static PresentationTextBody Read(P.TextBody? source)
    {
        var body = new PresentationTextBody();
        if (source is null) return body;
        var paragraphs = source.Elements<A.Paragraph>().ToArray();
        if (paragraphs.Length > MaxParagraphs)
            throw new CodecException("presentation_text_budget_exceeded", $"Presentation shape exceeds the {MaxParagraphs}-paragraph text budget.");
        var runCount = 0;
        foreach (var sourceParagraph in paragraphs)
        {
            var paragraph = new PresentationTextParagraph();
            var properties = sourceParagraph.ParagraphProperties;
            if (properties?.Level is not null) paragraph.Level = checked((uint)properties.Level.Value);
            if (properties?.Alignment?.Value is { } alignment && AlignmentName(alignment) is { Length: > 0 } alignmentName)
                paragraph.Alignment = alignmentName;
            foreach (var sourceRun in sourceParagraph.Elements<A.Run>())
            {
                runCount++;
                if (runCount > MaxRuns)
                    throw new CodecException("presentation_text_budget_exceeded", $"Presentation shape exceeds the {MaxRuns}-run text budget.");
                paragraph.Runs.Add(ReadRun(sourceRun));
            }
            body.Paragraphs.Add(paragraph);
        }
        return body;
    }

    internal static string Flatten(PresentationTextBody? body) =>
        body is null ? string.Empty : string.Join("\n", body.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))));

    internal static void NormalizeSemantics(PresentationShape shape)
    {
        var body = CanonicalBody(shape);
        shape.TextBody = body.Clone();
        shape.Text = Flatten(body);
    }

    internal static bool SupportsEditing(P.TextBody? body)
    {
        if (body is null) return true;
        var paragraphs = body.Elements<A.Paragraph>().ToArray();
        if (paragraphs.Length > MaxParagraphs) return false;
        var runs = 0;
        foreach (var paragraph in paragraphs)
        {
            if (paragraph.ChildElements.Any(child => child is not A.ParagraphProperties and not A.Run and not A.EndParagraphRunProperties)) return false;
            foreach (var run in paragraph.Elements<A.Run>())
            {
                runs++;
                if (runs > MaxRuns || run.ChildElements.Any(child => child is not A.RunProperties and not A.Text) || run.Elements<A.Text>().Count() != 1) return false;
            }
        }
        return true;
    }

    internal static void Validate(PresentationShape shape)
    {
        if (shape.TextBody is null) return;
        if (shape.TextBody.Paragraphs.Count > MaxParagraphs)
            throw new CodecException("presentation_text_budget_exceeded", $"Presentation shape exceeds the {MaxParagraphs}-paragraph text budget.");
        var runCount = 0;
        foreach (var paragraph in shape.TextBody.Paragraphs)
        {
            if (paragraph.HasLevel && paragraph.Level > 8)
                throw new CodecException("invalid_presentation_text", "Presentation paragraph level must be from 0 through 8.");
            if (paragraph.HasAlignment) ParseAlignment(paragraph.Alignment);
            foreach (var run in paragraph.Runs)
            {
                runCount++;
                if (runCount > MaxRuns)
                    throw new CodecException("presentation_text_budget_exceeded", $"Presentation shape exceeds the {MaxRuns}-run text budget.");
                if (run.HasFontSizePoints && (!(run.FontSizePoints > 0) || run.FontSizePoints > MaxFontSizePoints || !double.IsFinite(run.FontSizePoints)))
                    throw new CodecException("invalid_presentation_text", $"Presentation run font size must be finite and between 0 and {MaxFontSizePoints} points.");
                if (run.HasFontFamily && (string.IsNullOrWhiteSpace(run.FontFamily) || run.FontFamily.Length > 255))
                    throw new CodecException("invalid_presentation_text", "Presentation run font family must contain 1 through 255 characters.");
                if (run.HasColorRgb) PptxColor.Normalize(run.ColorRgb);
            }
        }
        _ = CanonicalBody(shape);
    }

    internal static P.TextBody Build(PresentationShape shape)
    {
        var body = new P.TextBody(new A.BodyProperties(), new A.ListStyle());
        var semantic = CanonicalBody(shape);
        foreach (var paragraph in semantic.Paragraphs) body.Append(BuildParagraph(paragraph));
        if (semantic.Paragraphs.Count == 0) body.Append(new A.Paragraph(new A.EndParagraphRunProperties { Language = "en-US" }));
        return body;
    }

    internal static void Apply(P.Shape shape, PresentationShape requested)
    {
        var semantic = CanonicalBody(requested);
        if (shape.TextBody is null)
        {
            if (semantic.Paragraphs.Count == 0 || (semantic.Paragraphs.Count == 1 && semantic.Paragraphs[0].Runs.Count == 0)) return;
            throw new CodecException("presentation_text_topology_changed", "Source-preserving PPTX export cannot add a text body to an imported shape.");
        }
        var paragraphs = shape.TextBody.Elements<A.Paragraph>().ToArray();
        if (paragraphs.Length != semantic.Paragraphs.Count)
            throw new CodecException("presentation_text_topology_changed", "Source-preserving PPTX export requires the original paragraph topology.");
        for (var paragraphIndex = 0; paragraphIndex < paragraphs.Length; paragraphIndex++)
        {
            var sourceParagraph = paragraphs[paragraphIndex];
            var requestedParagraph = semantic.Paragraphs[paragraphIndex];
            var runs = sourceParagraph.Elements<A.Run>().ToArray();
            if (runs.Length != requestedParagraph.Runs.Count)
                throw new CodecException("presentation_text_topology_changed", $"Source-preserving PPTX export requires paragraph {paragraphIndex + 1}'s original run topology.");
            ApplyParagraphProperties(sourceParagraph, requestedParagraph);
            for (var runIndex = 0; runIndex < runs.Length; runIndex++) ApplyRun(runs[runIndex], requestedParagraph.Runs[runIndex]);
        }
    }

    internal static void ScrubModeledContent(P.TextBody? body)
    {
        if (body is null) return;
        foreach (var paragraph in body.Elements<A.Paragraph>())
        {
            if (paragraph.ParagraphProperties is { } paragraphProperties)
            {
                paragraphProperties.Level = null;
                if (paragraphProperties.Alignment?.Value is { } alignment && AlignmentName(alignment).Length > 0)
                    paragraphProperties.Alignment = null;
            }
            foreach (var run in paragraph.Elements<A.Run>())
            {
                if (run.GetFirstChild<A.Text>() is { } text) text.Text = string.Empty;
                ScrubRunProperties(run.RunProperties);
            }
        }
    }

    private static PresentationTextBody CanonicalBody(PresentationShape shape)
    {
        if (shape.TextBody is not null)
        {
            if (shape.Text.Equals(Flatten(shape.TextBody), StringComparison.Ordinal)) return shape.TextBody;
            if (shape.TextBody.Paragraphs.Count != 1 || shape.TextBody.Paragraphs[0].Runs.Count > 1)
                throw new CodecException("presentation_text_mismatch", "Presentation shape text must equal structured text_body content for multi-paragraph or multi-run text.");
            var compatible = shape.TextBody.Clone();
            var compatibleParagraph = compatible.Paragraphs[0];
            if (compatibleParagraph.Runs.Count == 0) compatibleParagraph.Runs.Add(new PresentationTextRun());
            compatibleParagraph.Runs[0].Text = shape.Text;
            return compatible;
        }
        var body = new PresentationTextBody();
        var paragraph = new PresentationTextParagraph();
        if (shape.Text.Length > 0) paragraph.Runs.Add(new PresentationTextRun { Text = shape.Text });
        body.Paragraphs.Add(paragraph);
        return body;
    }

    private static PresentationTextRun ReadRun(A.Run source)
    {
        var run = new PresentationTextRun { Text = source.GetFirstChild<A.Text>()?.Text ?? string.Empty };
        var properties = source.RunProperties;
        if (properties?.Bold is not null) run.Bold = properties.Bold.Value;
        if (properties?.Italic is not null) run.Italic = properties.Italic.Value;
        if (properties?.FontSize is not null) run.FontSizePoints = properties.FontSize.Value / 100d;
        if (properties?.GetFirstChild<A.LatinFont>()?.Typeface?.Value is { Length: > 0 } typeface) run.FontFamily = typeface;
        if (PptxColor.SolidRgb(properties?.GetFirstChild<A.SolidFill>()) is { Length: > 0 } rgb) run.ColorRgb = rgb;
        return run;
    }

    private static A.Paragraph BuildParagraph(PresentationTextParagraph source)
    {
        var paragraph = new A.Paragraph();
        if (source.HasLevel || source.HasAlignment)
        {
            var properties = new A.ParagraphProperties();
            if (source.HasLevel) properties.Level = checked((int)source.Level);
            if (source.HasAlignment) properties.Alignment = ParseAlignment(source.Alignment);
            paragraph.Append(properties);
        }
        foreach (var run in source.Runs) paragraph.Append(BuildRun(run));
        paragraph.Append(new A.EndParagraphRunProperties { Language = "en-US" });
        return paragraph;
    }

    private static A.Run BuildRun(PresentationTextRun source)
    {
        var properties = new A.RunProperties { Language = "en-US" };
        ApplyRunProperties(properties, source);
        return new A.Run(properties, new A.Text(source.Text));
    }

    private static void ApplyParagraphProperties(A.Paragraph source, PresentationTextParagraph requested)
    {
        var properties = source.ParagraphProperties;
        if (properties is null && (requested.HasLevel || requested.HasAlignment))
        {
            properties = new A.ParagraphProperties();
            source.PrependChild(properties);
        }
        if (properties is null) return;
        properties.Level = requested.HasLevel ? checked((int)requested.Level) : null;
        if (requested.HasAlignment) properties.Alignment = ParseAlignment(requested.Alignment);
        else if (properties.Alignment?.Value is { } alignment && AlignmentName(alignment).Length > 0) properties.Alignment = null;
    }

    private static void ApplyRun(A.Run source, PresentationTextRun requested)
    {
        var properties = source.RunProperties;
        if (properties is null && HasStyle(requested))
        {
            properties = new A.RunProperties { Language = "en-US" };
            source.PrependChild(properties);
        }
        if (properties is not null) ApplyRunProperties(properties, requested);
        source.GetFirstChild<A.Text>()!.Text = requested.Text;
    }

    private static bool HasStyle(PresentationTextRun run) =>
        run.HasBold || run.HasItalic || run.HasFontSizePoints || run.HasFontFamily || run.HasColorRgb;

    private static void ApplyRunProperties(A.RunProperties properties, PresentationTextRun requested)
    {
        properties.Bold = requested.HasBold ? requested.Bold : null;
        properties.Italic = requested.HasItalic ? requested.Italic : null;
        properties.FontSize = requested.HasFontSizePoints ? checked((int)Math.Round(requested.FontSizePoints * 100)) : null;
        var latin = properties.GetFirstChild<A.LatinFont>();
        if (requested.HasFontFamily)
        {
            if (latin is null)
            {
                latin = new A.LatinFont();
                properties.Append(latin);
            }
            latin.Typeface = requested.FontFamily;
        }
        else latin?.Remove();
        var fill = properties.GetFirstChild<A.SolidFill>();
        if (requested.HasColorRgb)
        {
            fill?.Remove();
            properties.PrependChild(new A.SolidFill(new A.RgbColorModelHex { Val = PptxColor.Normalize(requested.ColorRgb) }));
        }
        else if (PptxColor.SolidRgb(fill).Length > 0) fill!.Remove();
    }

    private static void ScrubRunProperties(A.RunProperties? properties)
    {
        if (properties is null) return;
        properties.Bold = null;
        properties.Italic = null;
        properties.FontSize = null;
        properties.GetFirstChild<A.LatinFont>()?.Remove();
        var fill = properties.GetFirstChild<A.SolidFill>();
        if (PptxColor.SolidRgb(fill).Length > 0) fill!.Remove();
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
        _ => throw new CodecException("invalid_presentation_text", $"Unsupported Presentation paragraph alignment {value}."),
    };
}
