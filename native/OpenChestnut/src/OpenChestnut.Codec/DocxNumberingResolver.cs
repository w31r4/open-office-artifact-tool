using System.Xml.Linq;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Resolves the read-only Word numbering graph behind one paragraph. The
// resolver follows public ECMA-376 semantics: paragraph-style numPr supplies
// the numbering instance, w:lvl/w:pStyle selects the associated level, and a
// numStyleLink may redirect through a numbering style to the underlying
// abstract definition. Nothing in styles.xml or numbering.xml is materialized
// as an Open XML SDK root or mutated by this slice.
internal static class DocxNumberingResolver
{
    private const int MaxLinkDepth = 64;
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    internal static DocumentNumbering? Resolve(
        DocxPartContext context,
        int numberingId,
        int? directLevel,
        IReadOnlyList<string> paragraphStyleChain)
    {
        if (numberingId <= 0) return null;
        var numbering = context.NumberingDocument;
        if (numbering?.Root is null) return null;
        var instance = ResolveInstance(
            numbering,
            context.StylesDocument,
            numberingId,
            new HashSet<int>(),
            0);
        if (instance is null) return null;

        var levelIndex = directLevel ?? ResolveStyleLevel(instance.Levels, paragraphStyleChain);
        if (levelIndex is null or < 0 or > 8 || !instance.Levels.TryGetValue(levelIndex.Value, out var level))
            return null;
        return new DocumentNumbering
        {
            NumberingId = checked((uint)numberingId),
            Level = checked((uint)levelIndex.Value),
            AbstractNumberingId = checked((uint)instance.AbstractNumberingId),
            NumberFormat = level.NumberFormat,
            Start = checked((uint)level.Start),
            LevelText = level.LevelText,
            NumberingStyleId = instance.NumberingStyleId,
        };
    }

    private static int? ResolveStyleLevel(
        IReadOnlyDictionary<int, LevelDefinition> levels,
        IReadOnlyList<string> styleChain)
    {
        foreach (var styleId in styleChain)
        {
            var matches = levels.Values
                .Where(level => level.ParagraphStyleId.Equals(styleId, StringComparison.Ordinal))
                .Take(2)
                .ToArray();
            if (matches.Length > 1) return null;
            if (matches.Length == 1) return matches[0].Level;
        }
        return 0;
    }

    private static InstanceDefinition? ResolveInstance(
        XDocument numbering,
        XDocument? styles,
        int numberingId,
        HashSet<int> trail,
        int depth)
    {
        if (depth >= MaxLinkDepth || !trail.Add(numberingId)) return null;
        try
        {
            if (!TryUnique(numbering.Root?.Elements(W + "num") ?? [],
                    item => IntegerAttribute(item, W + "numId") == numberingId,
                    out var instance) || instance is null) return null;
            var abstractId = IntegerAttribute(instance.Element(W + "abstractNumId"), W + "val");
            if (abstractId is null or < 0) return null;
            if (!TryUnique(numbering.Root?.Elements(W + "abstractNum") ?? [],
                    item => IntegerAttribute(item, W + "abstractNumId") == abstractId.Value,
                    out var abstractNumbering) || abstractNumbering is null) return null;

            var styleLink = StringAttribute(abstractNumbering.Element(W + "styleLink"), W + "val");
            var numberingStyleLink = StringAttribute(abstractNumbering.Element(W + "numStyleLink"), W + "val");
            if ((styleLink?.Length ?? 0) > 253 || (numberingStyleLink?.Length ?? 0) > 253) return null;

            Dictionary<int, LevelDefinition> levels;
            var numberingStyleId = styleLink ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(numberingStyleLink))
            {
                var linkedNumberingId = ResolveNumberingStyle(styles, numberingStyleLink);
                if (linkedNumberingId is null or <= 0) return null;
                var linked = ResolveInstance(numbering, styles, linkedNumberingId.Value, trail, depth + 1);
                if (linked is null) return null;
                levels = linked.Levels.ToDictionary(item => item.Key, item => item.Value);
                numberingStyleId = numberingStyleLink;
            }
            else
            {
                var directLevels = ReadLevels(abstractNumbering.Elements(W + "lvl"));
                if (directLevels is null) return null;
                levels = directLevels;
            }

            var overrides = new HashSet<int>();
            foreach (var levelOverride in instance.Elements(W + "lvlOverride"))
            {
                var levelIndex = IntegerAttribute(levelOverride, W + "ilvl");
                if (levelIndex is null or < 0 or > 8 || !overrides.Add(levelIndex.Value)) return null;
                var nestedLevels = levelOverride.Elements(W + "lvl").Take(2).ToArray();
                if (nestedLevels.Length > 1) return null;
                if (nestedLevels.Length == 1)
                {
                    var nested = ReadLevel(nestedLevels[0], levelIndex.Value);
                    if (nested is null) return null;
                    levels[levelIndex.Value] = nested;
                    continue;
                }
                var startOverride = IntegerAttribute(levelOverride.Element(W + "startOverride"), W + "val");
                if (startOverride is null) continue;
                if (startOverride < 0 || !levels.TryGetValue(levelIndex.Value, out var current)) return null;
                levels[levelIndex.Value] = current with { Start = startOverride.Value };
            }

            return new InstanceDefinition(abstractId.Value, levels, numberingStyleId);
        }
        finally
        {
            trail.Remove(numberingId);
        }
    }

    private static int? ResolveNumberingStyle(XDocument? styles, string styleId)
    {
        if (styles?.Root is null) return null;
        if (!TryUnique(styles.Root.Elements(W + "style"),
                item => item.Attribute(W + "styleId")?.Value == styleId,
                out var style) || style is null) return null;
        if (!string.Equals(style.Attribute(W + "type")?.Value, "numbering", StringComparison.Ordinal)) return null;
        return IntegerAttribute(style.Element(W + "pPr")?.Element(W + "numPr")?.Element(W + "numId"), W + "val");
    }

    private static Dictionary<int, LevelDefinition>? ReadLevels(IEnumerable<XElement> source)
    {
        var levels = new Dictionary<int, LevelDefinition>();
        foreach (var element in source)
        {
            var level = ReadLevel(element);
            if (level is null || !levels.TryAdd(level.Level, level)) return null;
        }
        return levels;
    }

    private static LevelDefinition? ReadLevel(XElement element, int? expectedLevel = null)
    {
        var declaredLevel = IntegerAttribute(element, W + "ilvl");
        var level = expectedLevel ?? declaredLevel;
        if (level is null or < 0 or > 8 || declaredLevel is not null && declaredLevel != level) return null;
        var numberFormat = StringAttribute(element.Element(W + "numFmt"), W + "val") ?? string.Empty;
        var levelText = StringAttribute(element.Element(W + "lvlText"), W + "val") ?? string.Empty;
        var paragraphStyleId = StringAttribute(element.Element(W + "pStyle"), W + "val") ?? string.Empty;
        var start = IntegerAttribute(element.Element(W + "start"), W + "val") ?? 1;
        if (start < 0 || numberFormat.Length > 128 || levelText.Length > 1024 || paragraphStyleId.Length > 253)
            return null;
        return new LevelDefinition(level.Value, numberFormat, start, levelText, paragraphStyleId);
    }

    private static int? IntegerAttribute(XElement? element, XName name) =>
        int.TryParse(element?.Attribute(name)?.Value, out var value) ? value : null;

    private static string? StringAttribute(XElement? element, XName name) => element?.Attribute(name)?.Value;

    private static bool TryUnique(
        IEnumerable<XElement> elements,
        Func<XElement, bool> predicate,
        out XElement? result)
    {
        result = null;
        using var matches = elements.Where(predicate).Take(2).GetEnumerator();
        if (!matches.MoveNext()) return true;
        result = matches.Current;
        return !matches.MoveNext();
    }

    private sealed record InstanceDefinition(
        int AbstractNumberingId,
        Dictionary<int, LevelDefinition> Levels,
        string NumberingStyleId);

    private sealed record LevelDefinition(
        int Level,
        string NumberFormat,
        int Start,
        string LevelText,
        string ParagraphStyleId);
}
