using DocumentFormat.OpenXml.Packaging;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Builds only the direct, text-marker numbering graph represented by the
// public wire model. Picture bullets and style-linked graphs need additional
// asset/style wire data and therefore remain explicitly unsupported.
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
    internal sealed record LevelDefinition(int Level, string NumberFormat, int Start, string LevelText);

    internal static Plan CreatePlan(DocumentArtifact document)
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
            var definition = new LevelDefinition(level, source.NumberFormat, checked((int)source.Start), source.LevelText);
            if (levels.TryGetValue(level, out var existing) && existing != definition)
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

    internal static void Apply(MainDocumentPart mainPart, Plan plan)
    {
        if (plan.IsEmpty) return;
        var numbering = new W.Numbering();
        foreach (var definition in plan.Abstracts)
        {
            var abstractNumbering = new W.AbstractNum
            {
                AbstractNumberId = definition.Id,
            };
            abstractNumbering.Append(new W.MultiLevelType { Val = W.MultiLevelValues.Multilevel });
            foreach (var source in definition.Levels)
            {
                var level = new W.Level(
                    new W.StartNumberingValue { Val = source.Start },
                    new W.NumberingFormat { Val = new W.NumberFormatValues(source.NumberFormat) },
                    new W.LevelText { Val = source.LevelText },
                    new W.LevelJustification { Val = W.LevelJustificationValues.Left },
                    new W.PreviousParagraphProperties(
                        new W.Indentation
                        {
                            Left = ((source.Level + 1) * 720).ToString(),
                            Hanging = "360",
                        }))
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

        var part = mainPart.AddNewPart<NumberingDefinitionsPart>();
        part.Numbering = numbering;
        part.Numbering.Save();
    }

    private static CodecException Invalid(string message) => new("invalid_document_numbering", message);
    private static CodecException Unsupported(string message) => new("unsupported_document_features", message);
}
