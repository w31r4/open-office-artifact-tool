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
    private const int MaxInlines = 16_384;
    private const int MaxTabStops = 256;
    private const double MaxFontSizePoints = 768;

    internal static PresentationTextBody Read(P.TextBody? source, PptxSlideContext? slideContext = null)
    {
        var body = new PresentationTextBody();
        if (source is null) return body;
        var paragraphs = source.Elements<A.Paragraph>().ToArray();
        if (paragraphs.Length > MaxParagraphs)
            throw new CodecException("presentation_text_budget_exceeded", $"Presentation shape exceeds the {MaxParagraphs}-paragraph text budget.");
        var inlineCount = 0;
        foreach (var sourceParagraph in paragraphs)
        {
            var paragraph = new PresentationTextParagraph();
            var properties = sourceParagraph.ParagraphProperties;
            if (properties?.Level is not null) paragraph.Level = checked((uint)properties.Level.Value);
            if (properties?.Alignment?.Value is { } alignment && AlignmentName(alignment) is { Length: > 0 } alignmentName)
                paragraph.Alignment = alignmentName;
            PptxParagraphLayoutCodec.Read(paragraph, properties);
            PptxParagraphSpacingCodec.Read(paragraph, properties);
            PptxBulletCodec.Read(paragraph, properties, slideContext);
            PptxBulletStyleCodec.Read(paragraph, properties);
            PptxDefaultRunStyleCodec.Read(paragraph, properties);
            foreach (var sourceInline in ParagraphInlines(sourceParagraph))
            {
                inlineCount++;
                if (inlineCount > MaxInlines)
                    throw new CodecException("presentation_text_budget_exceeded", $"Presentation shape exceeds the {MaxInlines}-inline text budget.");
                paragraph.Runs.Add(ReadInline(sourceInline, slideContext));
            }
            ReadTabStops(paragraph, properties);
            body.Paragraphs.Add(paragraph);
        }
        return body;
    }

    internal static string Flatten(PresentationTextBody? body) =>
        body is null ? string.Empty : string.Join("\n", body.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(InlineText))));

    internal static void NormalizeSemantics(PresentationShape shape)
    {
        var body = CanonicalBody(shape);
        shape.TextBody = body.Clone();
        foreach (var run in shape.TextBody.Paragraphs.SelectMany(paragraph => paragraph.Runs))
        {
            // no_hyperlink is edit intent; native absence reimports as an unset
            // choice, so canonical semantic hashes collapse both forms.
            if (run.HyperlinkCase == PresentationTextRun.HyperlinkOneofCase.NoHyperlink) run.ClearHyperlink();
        }
        foreach (var paragraph in shape.TextBody.Paragraphs)
        {
            if (paragraph.HasNoTabStops) paragraph.ClearNoTabStops();
            if (paragraph.LeftMarginCase == PresentationTextParagraph.LeftMarginOneofCase.NoMarginLeft) paragraph.ClearLeftMargin();
            if (paragraph.IndentationCase == PresentationTextParagraph.IndentationOneofCase.NoIndent) paragraph.ClearIndentation();
            if (paragraph.LineSpacingCase == PresentationTextParagraph.LineSpacingOneofCase.NoLineSpacing) paragraph.ClearLineSpacing();
            if (paragraph.SpaceBeforeCase == PresentationTextParagraph.SpaceBeforeOneofCase.NoSpaceBefore) paragraph.ClearSpaceBefore();
            if (paragraph.SpaceAfterCase == PresentationTextParagraph.SpaceAfterOneofCase.NoSpaceAfter) paragraph.ClearSpaceAfter();
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.NoDefaultRunProperties) paragraph.ClearDefaultRunStyle();
        }
        shape.Text = Flatten(body);
    }

    internal static bool SupportsEditing(P.TextBody? body)
    {
        if (body is null) return true;
        var paragraphs = body.Elements<A.Paragraph>().ToArray();
        if (paragraphs.Length > MaxParagraphs) return false;
        var inlines = 0;
        foreach (var paragraph in paragraphs)
        {
            if (paragraph.ChildElements.Any(child => child is not A.ParagraphProperties and not A.Run and not A.Break and not A.Field and not A.EndParagraphRunProperties)) return false;
            if (!SupportsTabStops(paragraph.ParagraphProperties)) return false;
            if (!PptxParagraphLayoutCodec.Supports(paragraph.ParagraphProperties)) return false;
            if (!PptxParagraphSpacingCodec.Supports(paragraph.ParagraphProperties)) return false;
            if (!PptxDefaultRunStyleCodec.Supports(paragraph.ParagraphProperties)) return false;
            foreach (var inline in ParagraphInlines(paragraph))
            {
                inlines++;
                if (inlines > MaxInlines || !SupportsInline(inline)) return false;
            }
        }
        return true;
    }

    internal static void Validate(PresentationShape shape)
    {
        if (shape.TextBody is null) return;
        if (shape.TextBody.Paragraphs.Count > MaxParagraphs)
            throw new CodecException("presentation_text_budget_exceeded", $"Presentation shape exceeds the {MaxParagraphs}-paragraph text budget.");
        var inlineCount = 0;
        foreach (var paragraph in shape.TextBody.Paragraphs)
        {
            if (paragraph.HasLevel && paragraph.Level > 8)
                throw new CodecException("invalid_presentation_text", "Presentation paragraph level must be from 0 through 8.");
            if (paragraph.HasAlignment) ParseAlignment(paragraph.Alignment);
            PptxParagraphLayoutCodec.Validate(paragraph);
            PptxParagraphSpacingCodec.Validate(paragraph);
            PptxDefaultRunStyleCodec.Validate(paragraph);
            PptxBulletCodec.Validate(paragraph);
            PptxBulletStyleCodec.Validate(paragraph);
            ValidateTabStops(paragraph);
            foreach (var run in paragraph.Runs)
            {
                inlineCount++;
                if (inlineCount > MaxInlines)
                    throw new CodecException("presentation_text_budget_exceeded", $"Presentation shape exceeds the {MaxInlines}-inline text budget.");
                ValidateInlineContent(run);
                if (run.HasFontSizePoints && (!(run.FontSizePoints > 0) || run.FontSizePoints > MaxFontSizePoints || !double.IsFinite(run.FontSizePoints)))
                    throw new CodecException("invalid_presentation_text", $"Presentation run font size must be finite and between 0 and {MaxFontSizePoints} points.");
                if (run.HasFontFamily && (string.IsNullOrWhiteSpace(run.FontFamily) || run.FontFamily.Length > 255))
                    throw new CodecException("invalid_presentation_text", "Presentation run font family must contain 1 through 255 characters.");
                if (run.HasColorRgb) PptxColor.Normalize(run.ColorRgb);
                PptxHyperlinkCodec.Validate(run);
            }
        }
        _ = CanonicalBody(shape);
    }

    internal static P.TextBody Build(PresentationShape shape, PptxSlideContext? slideContext = null)
    {
        var body = new P.TextBody(new A.BodyProperties(), new A.ListStyle());
        var semantic = CanonicalBody(shape);
        foreach (var paragraph in semantic.Paragraphs) body.Append(BuildParagraph(paragraph, slideContext));
        if (semantic.Paragraphs.Count == 0) body.Append(new A.Paragraph(new A.EndParagraphRunProperties { Language = "en-US" }));
        return body;
    }

    internal static void Apply(P.Shape shape, PresentationShape requested, PptxSlideContext slideContext)
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
            var inlines = ParagraphInlines(sourceParagraph);
            if (inlines.Length != requestedParagraph.Runs.Count)
                throw new CodecException("presentation_text_topology_changed", $"Source-preserving PPTX export requires paragraph {paragraphIndex + 1}'s original inline topology.");
            ApplyParagraphProperties(sourceParagraph, requestedParagraph, slideContext);
            for (var inlineIndex = 0; inlineIndex < inlines.Length; inlineIndex++) ApplyInline(inlines[inlineIndex], requestedParagraph.Runs[inlineIndex], slideContext);
        }
    }

    internal static void ScrubModeledContent(P.TextBody? body, PptxSlideContext? slideContext = null)
    {
        if (body is null) return;
        foreach (var paragraph in body.Elements<A.Paragraph>())
        {
            if (paragraph.ParagraphProperties is { } paragraphProperties)
            {
                paragraphProperties.Level = null;
                if (paragraphProperties.Alignment?.Value is { } alignment && AlignmentName(alignment).Length > 0)
                    paragraphProperties.Alignment = null;
                PptxParagraphLayoutCodec.Scrub(paragraphProperties);
                PptxParagraphSpacingCodec.Scrub(paragraphProperties);
                PptxDefaultRunStyleCodec.Scrub(paragraphProperties);
                PptxBulletCodec.Scrub(paragraphProperties, slideContext);
                PptxBulletStyleCodec.Scrub(paragraphProperties);
                paragraphProperties.GetFirstChild<A.TabStopList>()?.Remove();
            }
            foreach (var inline in ParagraphInlines(paragraph))
            {
                if (inline.GetFirstChild<A.Text>() is { } text) text.Text = string.Empty;
                if (inline is A.Field field)
                {
                    field.Id = string.Empty;
                    field.Type = string.Empty;
                }
                ScrubRunProperties(InlineProperties(inline), slideContext);
            }
        }
    }

    private static PresentationTextBody CanonicalBody(PresentationShape shape)
    {
        if (shape.TextBody is not null)
        {
            if (shape.Text.Equals(Flatten(shape.TextBody), StringComparison.Ordinal)) return shape.TextBody;
            if (shape.TextBody.Paragraphs.Count != 1 || shape.TextBody.Paragraphs[0].Runs.Count > 1 || shape.TextBody.Paragraphs[0].Runs.FirstOrDefault()?.ContentCase is not (PresentationTextRun.ContentOneofCase.None or PresentationTextRun.ContentOneofCase.Text))
                throw new CodecException("presentation_text_mismatch", "Presentation shape text must equal structured text_body content for multi-paragraph or non-text inline content.");
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

    private static PresentationTextRun ReadInline(OpenXmlElement source, PptxSlideContext? slideContext)
    {
        var run = source switch
        {
            A.Run => new PresentationTextRun { Text = source.GetFirstChild<A.Text>()?.Text ?? string.Empty },
            A.Break => new PresentationTextRun { LineBreak = true },
            A.Field field => new PresentationTextRun
            {
                Field = new PresentationTextField
                {
                    Id = field.Id?.Value ?? string.Empty,
                    Type = field.Type?.Value ?? string.Empty,
                    Text = field.Text?.Text ?? string.Empty,
                },
            },
            _ => throw new CodecException("unsupported_presentation_text", $"Unsupported Presentation inline element {source.LocalName}."),
        };
        var properties = InlineProperties(source);
        if (properties?.Bold is not null) run.Bold = properties.Bold.Value;
        if (properties?.Italic is not null) run.Italic = properties.Italic.Value;
        if (properties?.FontSize is not null) run.FontSizePoints = properties.FontSize.Value / 100d;
        if (properties?.GetFirstChild<A.LatinFont>()?.Typeface?.Value is { Length: > 0 } typeface) run.FontFamily = typeface;
        if (PptxColor.SolidRgb(properties?.GetFirstChild<A.SolidFill>()) is { Length: > 0 } rgb) run.ColorRgb = rgb;
        PptxHyperlinkCodec.Read(run, properties, slideContext);
        return run;
    }

    private static A.Paragraph BuildParagraph(PresentationTextParagraph source, PptxSlideContext? slideContext)
    {
        var paragraph = new A.Paragraph();
        if (source.HasLevel || source.HasAlignment || PptxParagraphLayoutCodec.HasAuthoredLayout(source) || PptxParagraphSpacingCodec.HasAuthoredSpacing(source) || PptxBulletCodec.HasModeledBullet(source) || PptxBulletStyleCodec.HasModeledStyle(source) || PptxDefaultRunStyleCodec.HasAuthoredStyle(source) || source.TabStops.Count > 0)
        {
            var properties = new A.ParagraphProperties();
            if (source.HasLevel) properties.Level = checked((int)source.Level);
            if (source.HasAlignment) properties.Alignment = ParseAlignment(source.Alignment);
            PptxParagraphLayoutCodec.Append(properties, source);
            PptxParagraphSpacingCodec.Append(properties, source);
            PptxBulletStyleCodec.Append(properties, source);
            PptxBulletCodec.Append(properties, source, slideContext);
            AppendTabStops(properties, source);
            PptxDefaultRunStyleCodec.Append(properties, source);
            paragraph.Append(properties);
        }
        foreach (var run in source.Runs) paragraph.Append(BuildInline(run, slideContext));
        paragraph.Append(new A.EndParagraphRunProperties { Language = "en-US" });
        return paragraph;
    }

    private static OpenXmlElement BuildInline(PresentationTextRun source, PptxSlideContext? slideContext)
    {
        var properties = new A.RunProperties { Language = "en-US" };
        ApplyRunProperties(properties, source);
        PptxHyperlinkCodec.Append(properties, source, slideContext);
        return source.ContentCase switch
        {
            PresentationTextRun.ContentOneofCase.Text => new A.Run(properties, new A.Text(source.Text)),
            PresentationTextRun.ContentOneofCase.LineBreak => new A.Break(properties),
            PresentationTextRun.ContentOneofCase.Field => new A.Field(properties, new A.Text(source.Field.Text)) { Id = source.Field.Id, Type = source.Field.Type },
            _ => throw new CodecException("invalid_presentation_text", "Presentation inline must contain text, a line break, or a field."),
        };
    }

    private static void ApplyParagraphProperties(A.Paragraph source, PresentationTextParagraph requested, PptxSlideContext slideContext)
    {
        var properties = source.ParagraphProperties;
        if (properties is null && (requested.HasLevel || requested.HasAlignment || PptxParagraphLayoutCodec.HasAuthoredLayout(requested) || PptxParagraphSpacingCodec.HasAuthoredSpacing(requested) || PptxBulletCodec.HasModeledBullet(requested) || PptxBulletStyleCodec.HasModeledStyle(requested) || PptxDefaultRunStyleCodec.HasAuthoredStyle(requested) || requested.TabStops.Count > 0))
        {
            properties = new A.ParagraphProperties();
            source.PrependChild(properties);
        }
        if (properties is null) return;
        properties.Level = requested.HasLevel ? checked((int)requested.Level) : null;
        if (requested.HasAlignment) properties.Alignment = ParseAlignment(requested.Alignment);
        else if (properties.Alignment?.Value is { } alignment && AlignmentName(alignment).Length > 0) properties.Alignment = null;
        PptxParagraphLayoutCodec.Apply(properties, requested);
        PptxParagraphSpacingCodec.Apply(properties, requested);
        PptxBulletStyleCodec.Apply(properties, requested);
        PptxBulletCodec.Apply(properties, requested, slideContext);
        ApplyTabStops(properties, requested);
        PptxDefaultRunStyleCodec.Apply(properties, requested);
    }

    private static void ApplyInline(OpenXmlElement source, PresentationTextRun requested, PptxSlideContext slideContext)
    {
        if (!InlineKindMatches(source, requested))
            throw new CodecException("presentation_text_topology_changed", "Source-preserving PPTX export cannot change an inline between text, line-break, and field kinds.");
        var properties = InlineProperties(source);
        if (properties is null && (HasStyle(requested) || PptxHyperlinkCodec.HasModeledChoice(requested)))
        {
            properties = new A.RunProperties { Language = "en-US" };
            source.PrependChild(properties);
        }
        if (properties is not null)
        {
            ApplyRunProperties(properties, requested);
            PptxHyperlinkCodec.Apply(properties, requested, slideContext);
        }
        if (source is A.Run run) run.Text!.Text = requested.Text;
        else if (source is A.Field field)
        {
            field.Id = requested.Field.Id;
            field.Type = requested.Field.Type;
            field.Text!.Text = requested.Field.Text;
        }
    }

    private static OpenXmlElement[] ParagraphInlines(A.Paragraph paragraph) =>
        paragraph.ChildElements.Where(child => child is A.Run or A.Break or A.Field).ToArray();

    private static A.RunProperties? InlineProperties(OpenXmlElement source) => source switch
    {
        A.Run run => run.RunProperties,
        A.Break lineBreak => lineBreak.RunProperties,
        A.Field field => field.RunProperties,
        _ => null,
    };

    private static bool SupportsInline(OpenXmlElement source) => source switch
    {
        A.Run run => run.ChildElements.All(child => child is A.RunProperties or A.Text) &&
                     run.Elements<A.RunProperties>().Count() <= 1 && run.Elements<A.Text>().Count() == 1,
        A.Break lineBreak => lineBreak.ChildElements.All(child => child is A.RunProperties) &&
                             lineBreak.Elements<A.RunProperties>().Count() <= 1,
        A.Field field => field.ChildElements.All(child => child is A.RunProperties or A.ParagraphProperties or A.Text) &&
                         field.Elements<A.RunProperties>().Count() <= 1 && field.Elements<A.ParagraphProperties>().Count() <= 1 &&
                         field.Elements<A.Text>().Count() == 1 && ValidFieldId(field.Id?.Value) && ValidFieldType(field.Type?.Value),
        _ => false,
    };

    private static bool InlineKindMatches(OpenXmlElement source, PresentationTextRun requested) =>
        (source, requested.ContentCase) switch
        {
            (A.Run, PresentationTextRun.ContentOneofCase.Text) => true,
            (A.Break, PresentationTextRun.ContentOneofCase.LineBreak) => true,
            (A.Field, PresentationTextRun.ContentOneofCase.Field) => true,
            _ => false,
        };

    private static string InlineText(PresentationTextRun source) => source.ContentCase switch
    {
        PresentationTextRun.ContentOneofCase.Text => source.Text,
        PresentationTextRun.ContentOneofCase.LineBreak => "\n",
        PresentationTextRun.ContentOneofCase.Field => source.Field.Text,
        _ => string.Empty,
    };

    private static void ValidateInlineContent(PresentationTextRun source)
    {
        switch (source.ContentCase)
        {
            case PresentationTextRun.ContentOneofCase.Text:
                return;
            case PresentationTextRun.ContentOneofCase.LineBreak:
                if (!source.LineBreak) throw new CodecException("invalid_presentation_text", "Presentation line_break must be true when selected.");
                return;
            case PresentationTextRun.ContentOneofCase.Field:
                if (source.Field is null || !ValidFieldId(source.Field.Id))
                    throw new CodecException("invalid_presentation_text", "Presentation field id must be a brace-wrapped UUID.");
                if (!ValidFieldType(source.Field.Type))
                    throw new CodecException("invalid_presentation_text", "Presentation field type must contain 1 through 255 printable characters.");
                return;
            default:
                throw new CodecException("invalid_presentation_text", "Presentation inline must contain text, a line break, or a field.");
        }
    }

    private static bool ValidFieldId(string? value) => Guid.TryParseExact(value, "B", out _);

    private static bool ValidFieldType(string? value) =>
        !string.IsNullOrWhiteSpace(value) && value.Length <= 255 && !value.Any(char.IsControl);

    private static void ReadTabStops(PresentationTextParagraph target, A.ParagraphProperties? source)
    {
        var list = source?.GetFirstChild<A.TabStopList>();
        if (list is null) return;
        foreach (var tab in list.Elements<A.TabStop>())
        {
            if (tab.Position?.Value is not { } position || position < 0 || TabAlignmentName(tab.Alignment?.Value).Length == 0) continue;
            target.TabStops.Add(new PresentationTabStop { PositionEmu = position, Alignment = TabAlignmentName(tab.Alignment?.Value) });
        }
    }

    private static bool SupportsTabStops(A.ParagraphProperties? source)
    {
        if (source is null) return true;
        var lists = source.Elements<A.TabStopList>().ToArray();
        if (lists.Length > 1) return false;
        if (lists.Length == 0) return true;
        var tabs = lists[0].Elements<A.TabStop>().ToArray();
        if (tabs.Length > MaxTabStops || lists[0].ChildElements.Any(child => child is not A.TabStop)) return false;
        var previous = -1;
        foreach (var tab in tabs)
        {
            var position = tab.Position?.Value;
            if (position is null || position < 0 || position <= previous || TabAlignmentName(tab.Alignment?.Value).Length == 0) return false;
            previous = position.Value;
        }
        return true;
    }

    private static void ValidateTabStops(PresentationTextParagraph source)
    {
        if (source.HasNoTabStops)
        {
            if (!source.NoTabStops) throw new CodecException("invalid_presentation_text", "Presentation no_tab_stops must be true when selected.");
            if (source.TabStops.Count > 0) throw new CodecException("invalid_presentation_text", "Presentation tab_stops and no_tab_stops cannot both be selected.");
            return;
        }
        if (source.TabStops.Count == 0) return;
        if (source.TabStops.Count > MaxTabStops)
            throw new CodecException("presentation_text_budget_exceeded", $"Presentation paragraph exceeds the {MaxTabStops}-tab-stop budget.");
        long previous = -1;
        foreach (var tab in source.TabStops)
        {
            if (tab.PositionEmu < 0 || tab.PositionEmu > int.MaxValue || tab.PositionEmu <= previous)
                throw new CodecException("invalid_presentation_text", "Presentation tab stops must use strictly increasing positions in the signed 32-bit EMU range.");
            _ = ParseTabAlignment(tab.Alignment);
            previous = tab.PositionEmu;
        }
    }

    private static void AppendTabStops(A.ParagraphProperties target, PresentationTextParagraph source)
    {
        if (source.TabStops.Count == 0) return;
        var list = new A.TabStopList();
        foreach (var tab in source.TabStops)
            list.Append(new A.TabStop { Position = checked((int)tab.PositionEmu), Alignment = ParseTabAlignment(tab.Alignment) });
        target.AddChild(list, true);
    }

    private static void ApplyTabStops(A.ParagraphProperties target, PresentationTextParagraph source)
    {
        if (!source.HasNoTabStops && source.TabStops.Count == 0) return;
        var existing = target.Elements<A.TabStopList>().ToArray();
        if (existing.Length > 1 || !SupportsTabStops(target))
            throw new CodecException("unsupported_presentation_edit", "Source-preserving PPTX export cannot replace malformed or unmodeled tab stops.");
        foreach (var list in existing) list.Remove();
        AppendTabStops(target, source);
    }

    private static string TabAlignmentName(A.TextTabAlignmentValues? value)
    {
        if (value is null || value.Value == A.TextTabAlignmentValues.Left) return "left";
        if (value.Value == A.TextTabAlignmentValues.Center) return "center";
        if (value.Value == A.TextTabAlignmentValues.Right) return "right";
        if (value.Value == A.TextTabAlignmentValues.Decimal) return "decimal";
        return string.Empty;
    }

    private static A.TextTabAlignmentValues ParseTabAlignment(string value) => value switch
    {
        "left" => A.TextTabAlignmentValues.Left,
        "center" => A.TextTabAlignmentValues.Center,
        "right" => A.TextTabAlignmentValues.Right,
        "decimal" => A.TextTabAlignmentValues.Decimal,
        _ => throw new CodecException("invalid_presentation_text", $"Unsupported Presentation tab alignment {value}."),
    };

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

    private static void ScrubRunProperties(A.RunProperties? properties, PptxSlideContext? slideContext)
    {
        if (properties is null) return;
        properties.Bold = null;
        properties.Italic = null;
        properties.FontSize = null;
        properties.GetFirstChild<A.LatinFont>()?.Remove();
        var fill = properties.GetFirstChild<A.SolidFill>();
        if (PptxColor.SolidRgb(fill).Length > 0) fill!.Remove();
        PptxHyperlinkCodec.Scrub(properties, slideContext);
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
