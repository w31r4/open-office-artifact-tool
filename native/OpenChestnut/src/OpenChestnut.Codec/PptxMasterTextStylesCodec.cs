using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenChestnut.Codec;

// Owns the bounded p:txStyles paragraph-level defaults on one slide master.
// Master/layout identity and package topology remain in PptxCodec; this codec
// only mutates modeled lvl1pPr through lvl9pPr content and retains residual
// attributes, children, relationships, and unmodeled levels in place.
internal static class PptxMasterTextStylesCodec
{
    internal static PresentationMasterTextStyles Read(P.SlideMaster source, PptxPartContext? partContext)
    {
        var target = new PresentationMasterTextStyles();
        var styles = source.Elements<P.TextStyles>().FirstOrDefault();
        if (styles is null) return target;
        ReadLevels(target.TitleLevels, styles.Elements<P.TitleStyle>().FirstOrDefault(), partContext);
        ReadLevels(target.BodyLevels, styles.Elements<P.BodyStyle>().FirstOrDefault(), partContext);
        ReadLevels(target.OtherLevels, styles.Elements<P.OtherStyle>().FirstOrDefault(), partContext);
        return target;
    }

    internal static bool Supports(P.SlideMaster source)
    {
        var containers = source.Elements<P.TextStyles>().ToArray();
        if (containers.Length > 1) return false;
        var styles = containers.FirstOrDefault();
        return styles is null ||
               SupportsRoot(styles.Elements<P.TitleStyle>()) &&
               SupportsRoot(styles.Elements<P.BodyStyle>()) &&
               SupportsRoot(styles.Elements<P.OtherStyle>());
    }

    internal static void Validate(PresentationMasterTextStyles? source)
    {
        if (source is null) return;
        ValidateKind("title", source.TitleLevels, source.DeletedTitleLevels);
        ValidateKind("body", source.BodyLevels, source.DeletedBodyLevels);
        ValidateKind("other", source.OtherLevels, source.DeletedOtherLevels);
    }

    internal static void Build(P.SlideMaster target, PresentationMasterTextStyles? source, PptxPartContext? partContext)
    {
        if (source is null) return;
        Validate(source);
        var styles = target.GetFirstChild<P.TextStyles>() ?? target.AppendChild(new P.TextStyles());
        var title = EnsureRoot(styles, styles.GetFirstChild<P.TitleStyle>(), () => new P.TitleStyle());
        var body = EnsureRoot(styles, styles.GetFirstChild<P.BodyStyle>(), () => new P.BodyStyle());
        var other = EnsureRoot(styles, styles.GetFirstChild<P.OtherStyle>(), () => new P.OtherStyle());
        BuildRoot(title, source.TitleLevels, partContext);
        BuildRoot(body, source.BodyLevels, partContext);
        BuildRoot(other, source.OtherLevels, partContext);
    }

    internal static void Apply(P.SlideMaster target, PresentationMasterTextStyles source, PptxPartContext partContext)
    {
        Validate(source);
        var styles = target.GetFirstChild<P.TextStyles>();
        if (styles is null)
        {
            if (!HasEditIntent(source)) return;
            styles = target.AppendChild(new P.TextStyles());
        }
        ApplyRoot(styles, styles.GetFirstChild<P.TitleStyle>(), () => new P.TitleStyle(), source.TitleLevels, source.DeletedTitleLevels, partContext);
        ApplyRoot(styles, styles.GetFirstChild<P.BodyStyle>(), () => new P.BodyStyle(), source.BodyLevels, source.DeletedBodyLevels, partContext);
        ApplyRoot(styles, styles.GetFirstChild<P.OtherStyle>(), () => new P.OtherStyle(), source.OtherLevels, source.DeletedOtherLevels, partContext);
    }

    internal static void NormalizeSemantics(PresentationMasterTextStyles source)
    {
        NormalizeLevels(source.TitleLevels);
        NormalizeLevels(source.BodyLevels);
        NormalizeLevels(source.OtherLevels);
        source.DeletedTitleLevels.Clear();
        source.DeletedBodyLevels.Clear();
        source.DeletedOtherLevels.Clear();
    }

    internal static void ScrubModeledContent(P.SlideMaster source, PptxPartContext? partContext)
    {
        foreach (var styles in source.Elements<P.TextStyles>())
        foreach (var root in styles.ChildElements.OfType<OpenXmlCompositeElement>())
        foreach (var properties in LevelChildren(root).Values)
        {
            PptxParagraphPropertiesCodec.Scrub(properties, partContext, includeLevel: false);
            RemoveIfEmpty(properties);
        }
    }

    private static bool HasEditIntent(PresentationMasterTextStyles source) =>
        source.TitleLevels.Count > 0 || source.BodyLevels.Count > 0 || source.OtherLevels.Count > 0 ||
        source.DeletedTitleLevels.Count > 0 || source.DeletedBodyLevels.Count > 0 || source.DeletedOtherLevels.Count > 0;

    private static void ReadLevels(
        Google.Protobuf.Collections.RepeatedField<PresentationTextParagraph> target,
        OpenXmlCompositeElement? source,
        PptxPartContext? partContext)
    {
        if (source is null) return;
        foreach (var (level, properties) in LevelChildren(source).OrderBy(item => item.Key))
        {
            var modeled = new PresentationTextParagraph { Level = level };
            PptxParagraphPropertiesCodec.Read(modeled, properties, partContext, readLevel: false);
            if (PptxParagraphPropertiesCodec.HasModeledProperties(modeled)) target.Add(modeled);
        }
    }

    private static bool SupportsRoot<T>(IEnumerable<T> roots) where T : OpenXmlCompositeElement
    {
        var items = roots.ToArray();
        if (items.Length > 1) return false;
        if (items.Length == 0) return true;
        try
        {
            return LevelChildren(items[0]).Values.All(PptxParagraphPropertiesCodec.Supports);
        }
        catch (CodecException)
        {
            return false;
        }
    }

    private static void ValidateKind(
        string kind,
        Google.Protobuf.Collections.RepeatedField<PresentationTextParagraph> levels,
        Google.Protobuf.Collections.RepeatedField<uint> deletedLevels)
    {
        if (levels.Count > 9 || deletedLevels.Count > 9)
            throw Invalid($"Presentation master {kind} text style cannot contain more than nine levels.");
        var current = new HashSet<uint>();
        foreach (var level in levels)
        {
            PptxParagraphPropertiesCodec.Validate(level, requireLevel: true);
            if (!current.Add(level.Level)) throw Invalid($"Presentation master {kind} text style contains duplicate level {level.Level}.");
            if (level.Runs.Count > 0) throw Invalid("Presentation master text styles cannot contain text runs.");
            if (!PptxParagraphPropertiesCodec.HasModeledProperties(level))
                throw Invalid($"Presentation master {kind} text style level {level.Level} must contain at least one modeled property.");
        }
        var deleted = new HashSet<uint>();
        foreach (var level in deletedLevels)
        {
            if (level > 8) throw Invalid("Presentation master deleted text-style levels must be from 0 through 8.");
            if (!deleted.Add(level)) throw Invalid($"Presentation master {kind} text style contains duplicate deleted level {level}.");
            if (current.Contains(level)) throw Invalid($"Presentation master {kind} text style level {level} cannot be both present and deleted.");
        }
    }

    private static void BuildRoot(
        OpenXmlCompositeElement target,
        IEnumerable<PresentationTextParagraph> levels,
        PptxPartContext? partContext)
    {
        foreach (var level in levels.OrderBy(item => item.Level))
        {
            var properties = Create(level.Level);
            PptxParagraphPropertiesCodec.Append(properties, level, partContext, includeLevel: false);
            target.AddChild(properties, true);
        }
    }

    private static OpenXmlCompositeElement EnsureRoot(
        P.TextStyles owner,
        OpenXmlCompositeElement? existing,
        Func<OpenXmlCompositeElement> create)
    {
        if (existing is not null) return existing;
        var target = create();
        owner.AddChild(target, true);
        return target;
    }

    private static void ApplyRoot(
        P.TextStyles owner,
        OpenXmlCompositeElement? target,
        Func<OpenXmlCompositeElement> create,
        IEnumerable<PresentationTextParagraph> requestedLevels,
        IEnumerable<uint> deletedLevels,
        PptxPartContext partContext)
    {
        var requested = requestedLevels.ToDictionary(item => item.Level);
        var deleted = deletedLevels.ToHashSet();
        if (target is null)
        {
            if (requested.Count == 0) return;
            target = create();
            owner.AddChild(target, true);
        }
        var existing = LevelChildren(target);
        for (uint level = 0; level < 9; level++)
        {
            var current = existing.GetValueOrDefault(level);
            if (requested.TryGetValue(level, out var semantic))
            {
                if (current is null)
                {
                    current = Create(level);
                    target.AddChild(current, true);
                }
                PptxParagraphPropertiesCodec.Apply(current, semantic, partContext, includeLevel: false);
            }
            else if (deleted.Contains(level) && current is not null)
            {
                PptxParagraphPropertiesCodec.Scrub(current, partContext, includeLevel: false);
                RemoveIfEmpty(current);
            }
        }
    }

    private static void NormalizeLevels(Google.Protobuf.Collections.RepeatedField<PresentationTextParagraph> levels)
    {
        foreach (var level in levels) PptxTextCodec.NormalizeParagraphEditIntent(level);
        for (var index = levels.Count - 1; index >= 0; index--)
            if (!PptxParagraphPropertiesCodec.HasModeledProperties(levels[index])) levels.RemoveAt(index);
        var sorted = levels.OrderBy(item => item.Level).ToArray();
        levels.Clear();
        levels.Add(sorted);
    }

    private static Dictionary<uint, A.TextParagraphPropertiesType> LevelChildren(OpenXmlCompositeElement source)
    {
        var result = new Dictionary<uint, A.TextParagraphPropertiesType>();
        foreach (var child in source.ChildElements)
        {
            if (!TryLevel(child, out var level, out var properties)) continue;
            if (!result.TryAdd(level, properties)) throw Unsupported($"Source-preserving PPTX export cannot edit duplicate master text-style level {level}.");
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
        _ => throw Invalid("Presentation master text style level must be from 0 through 8."),
    };

    private static void RemoveIfEmpty(A.TextParagraphPropertiesType source)
    {
        if (source.GetAttributes().Count == 0 && source.ChildElements.Count == 0) source.Remove();
    }

    private static CodecException Invalid(string message) => new("invalid_presentation_master_style", message);
    private static CodecException Unsupported(string message) => new("unsupported_presentation_edit", message);
}
