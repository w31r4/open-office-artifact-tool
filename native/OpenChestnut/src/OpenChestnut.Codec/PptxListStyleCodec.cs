using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenChestnut.Codec;

// Owns a text body's a:lstStyle level topology. Each public list_styles entry
// is keyed by its required 0-based level and reuses the shared bounded
// CT_TextParagraphProperties implementation without writing a redundant lvl
// attribute to a:lvl1pPr through a:lvl9pPr.
internal static class PptxListStyleCodec
{
    internal static void Read(PresentationTextBody target, P.TextBody source, PptxPartContext? slideContext)
    {
        var listStyle = source.Elements<A.ListStyle>().FirstOrDefault();
        if (listStyle is null) return;
        var seen = new HashSet<uint>();
        foreach (var child in listStyle.ChildElements)
        {
            if (!TryLevel(child, out var level, out var properties) || !seen.Add(level)) continue;
            var modeled = new PresentationTextParagraph { Level = level };
            PptxParagraphPropertiesCodec.Read(modeled, properties, slideContext, readLevel: false);
            if (PptxParagraphPropertiesCodec.HasModeledProperties(modeled)) target.ListStyles.Add(modeled);
        }
    }

    internal static bool Supports(P.TextBody? source)
    {
        if (source is null) return true;
        var lists = source.Elements<A.ListStyle>().ToArray();
        if (lists.Length > 1) return false;
        if (lists.Length == 0) return true;
        var seen = new HashSet<uint>();
        foreach (var child in lists[0].ChildElements)
        {
            if (!TryLevel(child, out var level, out var properties)) continue;
            if (!seen.Add(level) || !PptxParagraphPropertiesCodec.Supports(properties)) return false;
        }
        return true;
    }

    internal static void Validate(PresentationTextBody source)
    {
        if (source.HasNoListStyles && !source.NoListStyles)
            throw Invalid("Presentation no_list_styles must be true when selected.");
        if (source.HasNoListStyles && source.ListStyles.Count > 0)
            throw Invalid("Presentation list_styles and no_list_styles cannot both be selected.");
        if (source.ListStyles.Count > 9)
            throw Invalid("Presentation text body cannot contain more than nine list styles.");
        var levels = new HashSet<uint>();
        foreach (var style in source.ListStyles)
        {
            PptxParagraphPropertiesCodec.Validate(style, requireLevel: true);
            if (!levels.Add(style.Level)) throw Invalid($"Presentation text body contains duplicate list style level {style.Level}.");
            if (style.Runs.Count > 0) throw Invalid("Presentation list styles cannot contain text runs.");
            if (!PptxParagraphPropertiesCodec.HasModeledProperties(style))
                throw Invalid($"Presentation list style level {style.Level} must contain at least one modeled property.");
        }
    }

    internal static void Build(A.ListStyle target, PresentationTextBody source, PptxPartContext? slideContext)
    {
        foreach (var style in source.ListStyles.OrderBy(item => item.Level))
        {
            var properties = Create(style.Level);
            PptxParagraphPropertiesCodec.Append(properties, style, slideContext, includeLevel: false);
            target.AddChild(properties, true);
        }
    }

    internal static void Apply(P.TextBody target, PresentationTextBody source, PptxPartContext slideContext)
    {
        var lists = target.Elements<A.ListStyle>().ToArray();
        if (lists.Length > 1) throw Unsupported("Source-preserving PPTX export cannot edit a malformed text-body list style.");
        var list = lists.FirstOrDefault();
        if (list is null)
        {
            if (source.ListStyles.Count == 0) return;
            list = new A.ListStyle();
            if (target.GetFirstChild<A.BodyProperties>() is { } bodyProperties) target.InsertAfter(list, bodyProperties);
            else target.PrependChild(list);
        }

        var existing = LevelChildren(list);
        var requested = source.ListStyles.ToDictionary(item => item.Level);
        for (uint level = 0; level < 9; level++)
        {
            var current = existing.GetValueOrDefault(level);
            if (requested.TryGetValue(level, out var style))
            {
                if (current is null)
                {
                    current = Create(level);
                    list.AddChild(current, true);
                }
                PptxParagraphPropertiesCodec.Apply(current, style, slideContext, includeLevel: false);
            }
            else if (current is not null)
            {
                PptxParagraphPropertiesCodec.Scrub(current, slideContext, includeLevel: false);
                RemoveIfEmpty(current);
            }
        }
    }

    internal static void Scrub(P.TextBody source, PptxPartContext? slideContext)
    {
        foreach (var list in source.Elements<A.ListStyle>())
        {
            foreach (var properties in LevelChildren(list).Values)
            {
                PptxParagraphPropertiesCodec.Scrub(properties, slideContext, includeLevel: false);
                RemoveIfEmpty(properties);
            }
        }
    }

    private static Dictionary<uint, A.TextParagraphPropertiesType> LevelChildren(A.ListStyle source)
    {
        var result = new Dictionary<uint, A.TextParagraphPropertiesType>();
        foreach (var child in source.ChildElements)
        {
            if (!TryLevel(child, out var level, out var properties)) continue;
            if (!result.TryAdd(level, properties)) throw Unsupported($"Source-preserving PPTX export cannot edit duplicate list style level {level}.");
        }
        return result;
    }

    private static bool TryLevel(OpenXmlElement source, out uint level, out A.TextParagraphPropertiesType properties)
    {
        switch (source)
        {
            case A.Level1ParagraphProperties value: level = 0; properties = value; return true;
            case A.Level2ParagraphProperties value: level = 1; properties = value; return true;
            case A.Level3ParagraphProperties value: level = 2; properties = value; return true;
            case A.Level4ParagraphProperties value: level = 3; properties = value; return true;
            case A.Level5ParagraphProperties value: level = 4; properties = value; return true;
            case A.Level6ParagraphProperties value: level = 5; properties = value; return true;
            case A.Level7ParagraphProperties value: level = 6; properties = value; return true;
            case A.Level8ParagraphProperties value: level = 7; properties = value; return true;
            case A.Level9ParagraphProperties value: level = 8; properties = value; return true;
            default: level = 0; properties = null!; return false;
        }
    }

    private static A.TextParagraphPropertiesType Create(uint level) => level switch
    {
        0 => new A.Level1ParagraphProperties(),
        1 => new A.Level2ParagraphProperties(),
        2 => new A.Level3ParagraphProperties(),
        3 => new A.Level4ParagraphProperties(),
        4 => new A.Level5ParagraphProperties(),
        5 => new A.Level6ParagraphProperties(),
        6 => new A.Level7ParagraphProperties(),
        7 => new A.Level8ParagraphProperties(),
        8 => new A.Level9ParagraphProperties(),
        _ => throw Invalid("Presentation list style level must be from 0 through 8."),
    };

    private static void RemoveIfEmpty(A.TextParagraphPropertiesType source)
    {
        if (source.GetAttributes().Count == 0 && source.ChildElements.Count == 0) source.Remove();
    }

    private static CodecException Invalid(string message) => new("invalid_presentation_text", message);
    private static CodecException Unsupported(string message) => new("unsupported_presentation_edit", message);
}
