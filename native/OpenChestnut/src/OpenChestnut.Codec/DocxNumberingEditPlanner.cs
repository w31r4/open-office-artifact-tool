using System.Xml.Linq;
using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Applies a deliberately bounded numbering-definition edit. The planner never
// changes a shared abstractNum. It materializes or updates an instance-local
// lvlOverride only when every top-level, directly assigned paragraph using the
// same numId/level requests the same definition and no other package part uses
// that numbering instance. Style-inherited and linked graphs remain read-only.
internal static class DocxNumberingEditPlanner
{
    private static readonly XNamespace Wml = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    internal sealed record Candidate(
        W.Paragraph Paragraph,
        DocumentNumbering Requested,
        DocumentNumbering Original,
        bool Editable);

    internal static void Apply(DocxPartContext context, IReadOnlyList<Candidate> candidates)
    {
        var edits = candidates.Where(candidate => !SameDefinition(candidate.Requested, candidate.Original)).ToArray();
        if (edits.Length == 0) return;

        var numbering = context.NumberingDocument ??
            throw Unsupported("Numbering-definition edits require a source Numbering part.");
        var authoredPictures = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var group in edits.GroupBy(candidate => Key(candidate.Original)))
        {
            var targetDefinitions = group.Select(candidate => Definition(candidate.Requested)).ToArray();
            if (targetDefinitions.Skip(1).Any(candidate => !SameDefinition(candidate, targetDefinitions[0])))
                throw Unsupported($"All paragraphs using numId {group.Key.NumberingId} level {group.Key.Level} must request one coherent numbering definition.");

            var affected = candidates.Where(candidate => Key(candidate.Original) == group.Key).ToArray();
            var target = targetDefinitions[0];
            if (affected.Any(candidate => !candidate.Editable || !SameIdentity(candidate.Requested, candidate.Original)))
                throw Unsupported($"Numbering-definition edits require every numId {group.Key.NumberingId} level {group.Key.Level} paragraph to keep its source identity and editable topology.");
            if (affected.Any(candidate => !SameDefinition(Definition(candidate.Requested), target)))
                throw Unsupported($"Numbering-definition edits must update every numId {group.Key.NumberingId} level {group.Key.Level} paragraph coherently.");
            if (affected.Any(candidate => !string.IsNullOrEmpty(candidate.Original.NumberingStyleId)))
                throw Unsupported("Numbering definitions reached through styleLink or numStyleLink remain source-bound.");
            if (affected.Any(candidate => !HasDirectAssignment(candidate.Paragraph, group.Key)))
                throw Unsupported("Numbering-definition edits currently require direct paragraph numPr assignments.");

            var directUses = context.Owner.Document?.Body?.Descendants<W.Paragraph>()
                .Where(paragraph => HasDirectAssignment(paragraph, group.Key))
                .ToArray() ?? [];
            if (directUses.Length != affected.Length || directUses.Any(paragraph => affected.All(candidate => !ReferenceEquals(candidate.Paragraph, paragraph))))
                throw Unsupported($"Numbering instance {group.Key.NumberingId} level {group.Key.Level} is also used by unmodeled or nested document content.");
            if (context.HasNumberingReferenceOutsideMainDocument(group.Key.NumberingId))
                throw Unsupported($"Numbering instance {group.Key.NumberingId} is referenced outside the modeled main-document paragraphs.");

            var original = Definition(group.First().Original);
            int? replacementPictureBulletId = null;
            if (!DocxPictureBulletCodec.SemanticKey(original.PictureBullet)
                    .Equals(DocxPictureBulletCodec.SemanticKey(target.PictureBullet), StringComparison.Ordinal))
            {
                if (original.PictureBullet is null || target.PictureBullet is null ||
                    original.PictureBullet.SourceCase != target.PictureBullet.SourceCase)
                    throw Unsupported($"Picture-bullet edits for numId {group.Key.NumberingId} level {group.Key.Level} must retain the source kind and cannot add or remove the native picture-bullet topology.");
                DocxPictureBulletCodec.Validate(target.PictureBullet, context.Images, $"DOCX numId {group.Key.NumberingId} level {group.Key.Level} picture bullet");
                var pictureKey = DocxPictureBulletCodec.SemanticKey(target.PictureBullet);
                if (!authoredPictures.TryGetValue(pictureKey, out var pictureBulletId))
                {
                    pictureBulletId = DocxPictureBulletCodec.AuthorSource(context, numbering, target.PictureBullet);
                    authoredPictures.Add(pictureKey, pictureBulletId);
                }
                replacementPictureBulletId = pictureBulletId;
            }

            ApplyGroup(numbering, group.Key, group.First().Original, target, replacementPictureBulletId);
        }
        context.SaveNumberingDocument();
    }

    private static void ApplyGroup(
        XDocument numbering,
        NumberingKey key,
        DocumentNumbering original,
        NumberingDefinition target,
        int? replacementPictureBulletId)
    {
        var root = numbering.Root ?? throw Unsupported("The source Numbering part has no root element.");
        var instance = Unique(root.Elements(Wml + "num"), element => IntegerAttribute(element, Wml + "numId") == key.NumberingId,
            $"Numbering instance {key.NumberingId} is missing or duplicated.");
        var abstractId = IntegerAttribute(instance.Element(Wml + "abstractNumId"), Wml + "val");
        if (abstractId != checked((int)original.AbstractNumberingId))
            throw Unsupported($"Numbering instance {key.NumberingId} no longer matches its source abstractNum identity.");
        var abstractNumbering = Unique(root.Elements(Wml + "abstractNum"), element => IntegerAttribute(element, Wml + "abstractNumId") == abstractId,
            $"Abstract numbering definition {abstractId} is missing or duplicated.");
        if (abstractNumbering.Element(Wml + "styleLink") is not null || abstractNumbering.Element(Wml + "numStyleLink") is not null)
            throw Unsupported("Linked abstract numbering definitions remain source-bound.");

        var overrides = instance.Elements(Wml + "lvlOverride")
            .Where(element => IntegerAttribute(element, Wml + "ilvl") == key.Level)
            .Take(2)
            .ToArray();
        if (overrides.Length > 1) throw Unsupported($"Numbering instance {key.NumberingId} has duplicate level {key.Level} overrides.");
        var levelOverride = overrides.SingleOrDefault();
        var nestedLevels = levelOverride?.Elements(Wml + "lvl").Take(2).ToArray() ?? [];
        if (nestedLevels.Length > 1) throw Unsupported($"Numbering instance {key.NumberingId} level {key.Level} has duplicate nested definitions.");
        if (nestedLevels.Length == 1)
        {
            SetDefinition(nestedLevels[0], key.Level, target, replacementPictureBulletId);
            return;
        }

        var startOverrides = levelOverride?.Elements(Wml + "startOverride").Take(2).ToArray() ?? [];
        if (startOverrides.Length > 1) throw Unsupported($"Numbering instance {key.NumberingId} level {key.Level} has duplicate start overrides.");
        var onlyStartChanged = original.NumberFormat == target.NumberFormat && original.LevelText == target.LevelText &&
            replacementPictureBulletId is null;
        if (levelOverride is not null && startOverrides.Length == 1 && onlyStartChanged)
        {
            startOverrides[0].SetAttributeValue(Wml + "val", target.Start);
            return;
        }
        if (levelOverride is not null)
            throw Unsupported($"Numbering instance {key.NumberingId} level {key.Level} has an unsupported partial override topology.");

        var sourceLevel = Unique(abstractNumbering.Elements(Wml + "lvl"), element => IntegerAttribute(element, Wml + "ilvl") == key.Level,
            $"Abstract numbering definition {abstractId} level {key.Level} is missing or duplicated.");
        var clonedLevel = new XElement(sourceLevel);
        SetDefinition(clonedLevel, key.Level, target, replacementPictureBulletId);
        instance.Add(new XElement(Wml + "lvlOverride",
            new XAttribute(Wml + "ilvl", key.Level),
            clonedLevel));
    }

    private static void SetDefinition(
        XElement level,
        int levelIndex,
        NumberingDefinition target,
        int? replacementPictureBulletId)
    {
        level.SetAttributeValue(Wml + "ilvl", levelIndex);
        SetLevelChild(level, "start", target.Start.ToString(), 0);
        SetLevelChild(level, "numFmt", target.NumberFormat, 1);
        SetLevelChild(level, "lvlText", target.LevelText, 6);
        if (replacementPictureBulletId is not null)
            SetLevelChild(level, "lvlPicBulletId", replacementPictureBulletId.Value.ToString(), 7);
    }

    private static void SetLevelChild(XElement level, string localName, string value, int order)
    {
        var matches = level.Elements(Wml + localName).Take(2).ToArray();
        if (matches.Length > 1) throw Unsupported($"Numbering level contains duplicate {localName} elements.");
        if (matches.Length == 1)
        {
            matches[0].SetAttributeValue(Wml + "val", value);
            return;
        }
        var child = new XElement(Wml + localName, new XAttribute(Wml + "val", value));
        var next = level.Elements().FirstOrDefault(element => LevelChildOrder(element.Name.LocalName) > order);
        if (next is null) level.Add(child);
        else next.AddBeforeSelf(child);
    }

    private static int LevelChildOrder(string localName) => localName switch
    {
        "start" => 0,
        "numFmt" => 1,
        "lvlRestart" => 2,
        "pStyle" => 3,
        "isLgl" => 4,
        "suff" => 5,
        "lvlText" => 6,
        "lvlPicBulletId" => 7,
        "legacy" => 8,
        "lvlJc" => 9,
        "pPr" => 10,
        "rPr" => 11,
        _ => int.MaxValue,
    };

    private static bool HasDirectAssignment(W.Paragraph paragraph, NumberingKey key)
    {
        var properties = paragraph.ParagraphProperties?.NumberingProperties;
        return properties?.NumberingId?.Val?.Value == key.NumberingId &&
               (properties.NumberingLevelReference?.Val?.Value ?? 0) == key.Level;
    }

    private static bool SameIdentity(DocumentNumbering left, DocumentNumbering right) =>
        left.NumberingId == right.NumberingId && left.Level == right.Level &&
        left.AbstractNumberingId == right.AbstractNumberingId &&
        left.NumberingStyleId.Equals(right.NumberingStyleId, StringComparison.Ordinal);

    private static bool SameDefinition(DocumentNumbering left, DocumentNumbering right) => SameDefinition(Definition(left), Definition(right));
    private static bool SameDefinition(NumberingDefinition left, NumberingDefinition right) =>
        left.NumberFormat.Equals(right.NumberFormat, StringComparison.Ordinal) && left.Start == right.Start &&
        left.LevelText.Equals(right.LevelText, StringComparison.Ordinal) &&
        DocxPictureBulletCodec.SemanticKey(left.PictureBullet).Equals(DocxPictureBulletCodec.SemanticKey(right.PictureBullet), StringComparison.Ordinal);
    private static NumberingKey Key(DocumentNumbering value) => new(checked((int)value.NumberingId), checked((int)value.Level));
    private static NumberingDefinition Definition(DocumentNumbering value) => new(value.NumberFormat, value.Start, value.LevelText, value.PictureBullet);

    private static XElement Unique(IEnumerable<XElement> source, Func<XElement, bool> predicate, string message)
    {
        var matches = source.Where(predicate).Take(2).ToArray();
        if (matches.Length != 1) throw Unsupported(message);
        return matches[0];
    }

    private static int? IntegerAttribute(XElement? element, XName name) =>
        int.TryParse(element?.Attribute(name)?.Value, out var value) ? value : null;

    private sealed record NumberingKey(int NumberingId, int Level);
    private sealed record NumberingDefinition(
        string NumberFormat,
        uint Start,
        string LevelText,
        DocumentPictureBullet? PictureBullet);
    private static CodecException Unsupported(string message) => new("unsupported_document_edit", message, "word/numbering.xml");
}
