using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Builds the direct numbering graph represented by the public wire model.
// Text markers and the canonical shared VML picture-bullet resource are
// source-built; style-linked graphs remain explicitly unsupported.
internal static class DocxDirectNumbering
{
    internal sealed record Plan(
        IReadOnlyList<AbstractDefinition> Abstracts,
        IReadOnlyList<InstanceDefinition> Instances)
    {
        internal bool IsEmpty => Instances.Count == 0;
    }

    internal sealed record AbstractDefinition(int Id, IReadOnlyList<LevelDefinition> Levels);
    internal sealed record InstanceDefinition(int Id, int AbstractId);
    internal sealed record LevelDefinition(
        int Level,
        string NumberFormat,
        int Start,
        string LevelText,
        DocumentPictureBullet? PictureBullet);

    internal static Plan CreatePlan(DocumentArtifact document, DocxImageAssetCatalog images)
    {
        var instanceAbstracts = new Dictionary<int, int>();
        var levelsByAbstract = new Dictionary<int, Dictionary<int, LevelDefinition>>();

        foreach (var block in document.Blocks.Where(block =>
                     block.ContentCase == DocumentBlock.ContentOneofCase.Paragraph &&
                     block.Paragraph.Numbering is not null))
        {
            var source = block.Paragraph.Numbering;
            DocxNumberedParagraphCodec.Validate(block.Paragraph);
            if (!string.IsNullOrWhiteSpace(source.NumberingStyleId))
                throw Unsupported($"Direct DOCX numbering for block {block.Id} cannot author a style-linked numbering graph.");
            if (source.Start == 0)
                throw Invalid($"Direct DOCX numbering for block {block.Id} requires a positive start value.");
            if (string.IsNullOrWhiteSpace(source.NumberFormat))
                throw Invalid($"Direct DOCX numbering for block {block.Id} requires a non-empty number_format.");
            if (string.IsNullOrEmpty(source.LevelText))
                throw Invalid($"Direct DOCX numbering for block {block.Id} requires non-empty level_text.");
            DocxPictureBulletCodec.Validate(source.PictureBullet, images, $"Direct DOCX numbering for block {block.Id}");
            if (source.PictureBullet is not null &&
                (!source.NumberFormat.Equals("bullet", StringComparison.Ordinal) || source.LevelText.EnumerateRunes().Count() != 1))
                throw Invalid($"Direct DOCX picture bullet for block {block.Id} requires bullet number_format and exactly one level_text character.");

            var numberingId = checked((int)source.NumberingId);
            var abstractId = checked((int)source.AbstractNumberingId);
            var level = checked((int)source.Level);
            if (instanceAbstracts.TryGetValue(numberingId, out var existingAbstractId) && existingAbstractId != abstractId)
                throw Invalid($"DOCX numbering instance {numberingId} references conflicting abstract numbering IDs {existingAbstractId} and {abstractId}.");
            instanceAbstracts[numberingId] = abstractId;

            if (!levelsByAbstract.TryGetValue(abstractId, out var levels))
            {
                levels = new Dictionary<int, LevelDefinition>();
                levelsByAbstract.Add(abstractId, levels);
            }
            var definition = new LevelDefinition(level, source.NumberFormat, checked((int)source.Start), source.LevelText, source.PictureBullet?.Clone());
            if (levels.TryGetValue(level, out var existing) && !SameDefinition(existing, definition))
                throw Invalid($"DOCX abstract numbering {abstractId} level {level} has conflicting definitions.");
            levels[level] = definition;
        }

        return new Plan(
            levelsByAbstract
                .OrderBy(item => item.Key)
                .Select(item => new AbstractDefinition(item.Key, item.Value.Values.OrderBy(level => level.Level).ToArray()))
                .ToArray(),
            instanceAbstracts
                .OrderBy(item => item.Key)
                .Select(item => new InstanceDefinition(item.Key, item.Value))
                .ToArray());
    }

    internal static void Apply(MainDocumentPart mainPart, Plan plan, DocxImageAssetCatalog images)
    {
        if (plan.IsEmpty) return;
        var part = mainPart.AddNewPart<NumberingDefinitionsPart>();
        var numbering = new W.Numbering();
        var pictureIds = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var source in plan.Abstracts
                     .SelectMany(definition => definition.Levels)
                     .Where(level => level.PictureBullet is not null))
        {
            var key = DocxPictureBulletCodec.SemanticKey(source.PictureBullet);
            if (pictureIds.ContainsKey(key)) continue;
            var id = pictureIds.Count;
            pictureIds.Add(key, id);
            numbering.Append(DocxPictureBulletCodec.Author(part, id, source.PictureBullet!, images));
        }
        foreach (var definition in plan.Abstracts)
        {
            var abstractNumbering = new W.AbstractNum
            {
                AbstractNumberId = definition.Id,
            };
            abstractNumbering.Append(new W.MultiLevelType { Val = W.MultiLevelValues.Multilevel });
            foreach (var source in definition.Levels)
            {
                var children = new List<OpenXmlElement>
                {
                    new W.StartNumberingValue { Val = source.Start },
                    new W.NumberingFormat { Val = new W.NumberFormatValues(source.NumberFormat) },
                    new W.LevelText { Val = source.LevelText },
                };
                if (source.PictureBullet is not null)
                    children.Add(new W.LevelPictureBulletId { Val = pictureIds[DocxPictureBulletCodec.SemanticKey(source.PictureBullet)] });
                children.Add(new W.LevelJustification { Val = W.LevelJustificationValues.Left });
                children.Add(new W.PreviousParagraphProperties(
                    new W.Indentation
                    {
                        Left = ((source.Level + 1) * 720).ToString(),
                        Hanging = "360",
                    }));
                var level = new W.Level(children)
                {
                    LevelIndex = source.Level,
                };
                if (source.NumberFormat.Equals("bullet", StringComparison.Ordinal))
                    level.Append(new W.NumberingSymbolRunProperties(
                        new W.RunFonts { Ascii = "Symbol", HighAnsi = "Symbol" }));
                abstractNumbering.Append(level);
            }
            numbering.Append(abstractNumbering);
        }
        foreach (var source in plan.Instances)
        {
            numbering.Append(new W.NumberingInstance(
                new W.AbstractNumId { Val = source.AbstractId })
            {
                NumberID = source.Id,
            });
        }
        part.Numbering = numbering;
        part.Numbering.Save();
    }

    private static bool SameDefinition(LevelDefinition left, LevelDefinition right) =>
        left.Level == right.Level && left.NumberFormat.Equals(right.NumberFormat, StringComparison.Ordinal) &&
        left.Start == right.Start && left.LevelText.Equals(right.LevelText, StringComparison.Ordinal) &&
        DocxPictureBulletCodec.SemanticKey(left.PictureBullet).Equals(DocxPictureBulletCodec.SemanticKey(right.PictureBullet), StringComparison.Ordinal);

    private static CodecException Invalid(string message) => new("invalid_document_numbering", message);
    private static CodecException Unsupported(string message) => new("unsupported_document_features", message);
}
