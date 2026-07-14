using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Models a direct w:numPr assignment and its resolved numbering level while
// keeping the shared numbering graph source-bound. The initial editable shape
// is deliberately one run/one text node: only that text may change.
internal static class DocxNumberedParagraphCodec
{
    internal static bool TryRead(
        W.Paragraph paragraph,
        DocxPartContext context,
        out DocumentParagraph artifact,
        out bool editable)
    {
        artifact = new DocumentParagraph();
        editable = false;
        var properties = paragraph.ParagraphProperties;
        var numberingProperties = properties?.NumberingProperties;
        var numberingId = numberingProperties?.NumberingId?.Val?.Value;
        if (numberingProperties is null || numberingId is null || numberingId <= 0) return false;

        var level = numberingProperties.NumberingLevelReference?.Val?.Value ?? 0;
        artifact.Text = string.Concat(paragraph.Descendants<W.Text>().Select(text => text.Text));
        var resolved = Resolve(context, numberingId.Value, level);
        artifact.Numbering = resolved ?? new DocumentNumbering
        {
            NumberingId = checked((uint)numberingId.Value),
            Level = checked((uint)Math.Max(0, level)),
        };
        if (paragraph.Elements<W.Run>().Count() == 1)
            artifact.Runs.Add(ReadRun(paragraph.Elements<W.Run>().Single()));
        editable = resolved is not null && IsEditable(paragraph);
        return true;
    }

    internal static void Apply(W.Paragraph paragraph, DocumentParagraph requested, DocumentParagraph original)
    {
        Validate(requested);
        if (!IsEditable(paragraph)) throw Unsupported("Source-preserving DOCX export cannot edit this numbered paragraph topology.");
        if (!SameNumbering(requested.Numbering, original.Numbering))
            throw Unsupported("Numbering identity, level, and shared definition metadata are source-bound in this codec slice.");
        if (requested.Runs.Count != 1 || requested.Text != requested.Runs[0].Text)
            throw Unsupported("Source-preserving numbered paragraphs require one modeled run whose text matches the paragraph text.");

        var text = paragraph.Descendants<W.Text>().Single();
        text.Text = requested.Text;
        text.Space = requested.Text.Length != requested.Text.Trim().Length ? SpaceProcessingModeValues.Preserve : null;
    }

    internal static string ResidualHash(W.Paragraph paragraph)
    {
        var clone = (W.Paragraph)paragraph.CloneNode(true);
        foreach (var text in clone.Descendants<W.Text>())
        {
            text.Text = string.Empty;
            text.Space = null;
        }
        return Hash(Encoding.UTF8.GetBytes(clone.OuterXml));
    }

    internal static void Validate(DocumentParagraph paragraph)
    {
        if (paragraph.Numbering is null) throw Invalid("Numbered paragraph metadata is missing.");
        if (paragraph.Numbering.NumberingId == 0) throw Invalid("Numbered paragraph numbering_id must be greater than zero.");
        if (paragraph.Numbering.Level > 8) throw Invalid("Numbered paragraph level must be between 0 and 8.");
        if (paragraph.Numbering.NumberFormat.Length > 128) throw Invalid("Numbered paragraph number_format exceeds 128 characters.");
        if (paragraph.Numbering.LevelText.Length > 1024) throw Invalid("Numbered paragraph level_text exceeds 1024 characters.");
        if (paragraph.Text.Length > 1_000_000) throw Invalid("Numbered paragraph text exceeds 1,000,000 characters.");
    }

    private static DocumentNumbering? Resolve(DocxPartContext context, int numberingId, int levelIndex)
    {
        if (levelIndex is < 0 or > 8) return null;
        var part = context.Owner.NumberingDefinitionsPart;
        if (part is null) return null;
        XDocument document;
        using (var stream = part.GetStream(FileMode.Open, FileAccess.Read))
        using (var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
        }))
            document = XDocument.Load(reader, LoadOptions.None);

        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        static int? IntegerAttribute(XElement? element, XName name) =>
            int.TryParse(element?.Attribute(name)?.Value, out var value) ? value : null;
        var instance = document.Root?.Elements(w + "num")
            .SingleOrDefault(item => IntegerAttribute(item, w + "numId") == numberingId);
        var abstractId = IntegerAttribute(instance?.Element(w + "abstractNumId"), w + "val");
        if (abstractId is null or < 0) return null;
        var abstractNumbering = document.Root?.Elements(w + "abstractNum")
            .SingleOrDefault(item => IntegerAttribute(item, w + "abstractNumId") == abstractId.Value);
        if (abstractNumbering is null) return null;
        var levelOverride = instance?.Elements(w + "lvlOverride")
            .SingleOrDefault(item => IntegerAttribute(item, w + "ilvl") == levelIndex);
        var level = levelOverride?.Element(w + "lvl") ?? abstractNumbering.Elements(w + "lvl")
            .SingleOrDefault(item => IntegerAttribute(item, w + "ilvl") == levelIndex);
        if (level is null) return null;
        var format = level.Element(w + "numFmt")?.Attribute(w + "val")?.Value ?? string.Empty;
        var levelText = level.Element(w + "lvlText")?.Attribute(w + "val")?.Value ?? string.Empty;
        var start = IntegerAttribute(levelOverride?.Element(w + "startOverride"), w + "val") ??
                    IntegerAttribute(level.Element(w + "start"), w + "val") ?? 1;
        if (start < 0) return null;
        return new DocumentNumbering
        {
            NumberingId = checked((uint)numberingId),
            Level = checked((uint)levelIndex),
            AbstractNumberingId = checked((uint)abstractId.Value),
            NumberFormat = format,
            Start = checked((uint)start),
            LevelText = levelText,
        };
    }

    private static DocumentRun ReadRun(W.Run run)
    {
        var properties = run.RunProperties;
        return new DocumentRun
        {
            Text = string.Concat(run.Descendants<W.Text>().Select(text => text.Text)),
            StyleId = properties?.RunStyle?.Val?.Value ?? string.Empty,
            Bold = IsOn(properties?.Bold),
            Italic = IsOn(properties?.Italic),
            Underline = properties?.Underline?.Val?.Value is { } underline && !underline.Equals(W.UnderlineValues.None),
        };
    }

    private static bool IsEditable(W.Paragraph paragraph)
    {
        if (paragraph.ChildElements.Any(child => child is not W.ParagraphProperties and not W.Run)) return false;
        var properties = paragraph.ParagraphProperties;
        if (properties is null || properties.ChildElements.Any(child => child is not W.ParagraphStyleId and not W.NumberingProperties)) return false;
        var numbering = properties.NumberingProperties;
        if (numbering is null || numbering.ChildElements.Any(child => child is not W.NumberingLevelReference and not W.NumberingId)) return false;
        if (numbering.NumberingId?.Val?.Value is null or <= 0) return false;
        var runs = paragraph.Elements<W.Run>().ToArray();
        if (runs.Length != 1) return false;
        var run = runs[0];
        if (run.ChildElements.Any(child => child is not W.RunProperties and not W.Text)) return false;
        if (run.Elements<W.Text>().Count() != 1) return false;
        return run.RunProperties?.ChildElements.All(child =>
            child is W.RunStyle or W.Bold or W.Italic or W.Underline) ?? true;
    }

    private static bool SameNumbering(DocumentNumbering? left, DocumentNumbering? right) =>
        left is not null && right is not null && left.Equals(right);

    private static bool IsOn(W.OnOffType? value) => value is not null && value.Val?.Value != false;
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static CodecException Invalid(string message) => new("invalid_document_numbering", message);
    private static CodecException Unsupported(string message) => new("unsupported_document_edit", message, "word/document.xml");
}
