using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

internal sealed class DocxHeaderFooterPlan
{
    private readonly Dictionary<uint, List<OpenXmlElement>> _references = [];
    private readonly Dictionary<uint, bool> _differentFirstPage = [];

    internal void AddReference(uint sectionIndex, OpenXmlElement reference)
    {
        if (!_references.TryGetValue(sectionIndex, out var references))
            _references.Add(sectionIndex, references = []);
        references.Add(reference);
    }

    internal void SetDifferentFirstPage(uint sectionIndex, bool value) =>
        _differentFirstPage[sectionIndex] = value;

    internal IReadOnlyList<OpenXmlElement> References(uint sectionIndex) =>
        _references.TryGetValue(sectionIndex, out var references) ? references : [];

    internal bool DifferentFirstPage(uint sectionIndex) =>
        _differentFirstPage.GetValueOrDefault(sectionIndex);
}

internal static class DocxHeaderFooterCodec
{
    private sealed record Group(bool Header, uint SectionIndex, DocumentHeaderFooterReference Reference, IReadOnlyList<DocumentHeaderFooter> Blocks);

    internal static void Read(
        MainDocumentPart mainPart,
        W.Body body,
        DocumentArtifact document,
        ICollection<Diagnostic> diagnostics)
    {
        var sections = BoundarySections(body).ToArray();
        for (var sectionIndex = 0; sectionIndex < sections.Length; sectionIndex++)
        {
            var properties = sections[sectionIndex];
            if (properties.GetFirstChild<W.TitlePage>() is not null)
                document.SectionSettings.Add(new DocumentSectionSettings
                {
                    SectionIndex = checked((uint)sectionIndex),
                    DifferentFirstPage = true,
                });
            ReadReferences(mainPart, properties, checked((uint)sectionIndex), header: true, document, diagnostics);
            ReadReferences(mainPart, properties, checked((uint)sectionIndex), header: false, document, diagnostics);
        }
    }

    internal static DocxHeaderFooterPlan Author(MainDocumentPart mainPart, DocumentArtifact document)
    {
        Validate(document);
        var sectionCount = SectionCount(document);
        var plan = new DocxHeaderFooterPlan();
        foreach (var settings in document.SectionSettings)
            plan.SetDifferentFirstPage(settings.SectionIndex,
                settings.HasDifferentFirstPage && settings.DifferentFirstPage);

        var groups = Groups(document, sectionCount).ToArray();
        foreach (var group in groups)
        {
            var activeFirst = group.Reference == DocumentHeaderFooterReference.First &&
                group.Blocks.Any(block => !block.HasVariantActive || block.VariantActive);
            if (activeFirst && document.SectionSettings.All(item => item.SectionIndex != group.SectionIndex))
                plan.SetDifferentFirstPage(group.SectionIndex, true);

            if (group.Header)
            {
                var part = mainPart.AddNewPart<HeaderPart>();
                part.Header = new W.Header(group.Blocks.Select(BuildParagraph));
                part.Header.Save();
                plan.AddReference(group.SectionIndex, new W.HeaderReference
                {
                    Type = ToNativeReference(group.Reference),
                    Id = mainPart.GetIdOfPart(part),
                });
            }
            else
            {
                var part = mainPart.AddNewPart<FooterPart>();
                part.Footer = new W.Footer(group.Blocks.Select(BuildParagraph));
                part.Footer.Save();
                plan.AddReference(group.SectionIndex, new W.FooterReference
                {
                    Type = ToNativeReference(group.Reference),
                    Id = mainPart.GetIdOfPart(part),
                });
            }
        }

        return plan;
    }

    internal static void AssertSourceUnchanged(
        MainDocumentPart mainPart,
        W.Body body,
        DocumentArtifact requested)
    {
        var source = new DocumentArtifact();
        var ignored = new List<Diagnostic>();
        DocxSettingsCodec.Read(mainPart, source);
        Read(mainPart, body, source, ignored);
        if (!SequenceEqual(source.SectionSettings, requested.SectionSettings) ||
            !SequenceEqual(source.Headers, requested.Headers) ||
            !SequenceEqual(source.Footers, requested.Footers))
            throw new CodecException(
                "unsupported_document_header_footer_edit",
                "Source-preserving DOCX export requires the imported header/footer topology, text, fields, and section activation settings to remain unchanged.",
                "word/document.xml");
    }

    internal static void Validate(DocumentArtifact document)
    {
        var sectionCount = SectionCount(document);
        var settings = new HashSet<uint>();
        foreach (var item in document.SectionSettings)
        {
            if (item.SectionIndex >= sectionCount || !settings.Add(item.SectionIndex))
                throw new CodecException("invalid_document_section_settings", $"Document section settings index {item.SectionIndex} is duplicate or outside 0 through {sectionCount - 1}.");
        }

        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var (kind, item) in document.Headers.Select(item => ("header", item))
                     .Concat(document.Footers.Select(item => ("footer", item))))
        {
            if (string.IsNullOrWhiteSpace(item.Id) || item.Id.Length > 512 || item.Id.Any(char.IsControl) || !ids.Add(item.Id))
                throw new CodecException("invalid_document_header_footer", $"Document {kind} IDs must be unique and contain 1 through 512 characters without controls.");
            if (item.Reference == DocumentHeaderFooterReference.Unspecified)
                throw new CodecException("invalid_document_header_footer", $"Document {kind} {item.Id} requires default, first, or even reference type.");
            var sectionIndex = item.HasSectionIndex ? item.SectionIndex : sectionCount - 1;
            if (sectionIndex >= sectionCount)
                throw new CodecException("invalid_document_header_footer", $"Document {kind} {item.Id} section index is outside 0 through {sectionCount - 1}.");
            if (item.Text.Length > 1_000_000 || item.Text.Any(character => character is '\0'))
                throw new CodecException("invalid_document_header_footer", $"Document {kind} {item.Id} text is invalid or too long.");
            if (!string.IsNullOrWhiteSpace(item.FieldInstruction))
                DocxFieldCodec.Validate(new DocumentField { Instruction = item.FieldInstruction, Display = item.Text });
        }
    }

    internal static uint SectionCount(DocumentArtifact document) =>
        checked((uint)document.Blocks.Count(block => block.ContentCase == DocumentBlock.ContentOneofCase.Section) + 1U);

    private static IEnumerable<W.SectionProperties> BoundarySections(W.Body body)
    {
        foreach (var paragraph in body.Elements<W.Paragraph>())
            if (paragraph.ParagraphProperties?.SectionProperties is { } boundary) yield return boundary;
        if (body.Elements<W.SectionProperties>().LastOrDefault() is { } final) yield return final;
        else yield return new W.SectionProperties();
    }

    private static void ReadReferences(
        MainDocumentPart mainPart,
        W.SectionProperties properties,
        uint sectionIndex,
        bool header,
        DocumentArtifact document,
        ICollection<Diagnostic> diagnostics)
    {
        var references = header
            ? properties.Elements<W.HeaderReference>().Cast<OpenXmlElement>()
            : properties.Elements<W.FooterReference>().Cast<OpenXmlElement>();
        foreach (var reference in references)
        {
            var relationshipId = header
                ? ((W.HeaderReference)reference).Id?.Value
                : ((W.FooterReference)reference).Id?.Value;
            var type = header
                ? ((W.HeaderReference)reference).Type?.Value
                : ((W.FooterReference)reference).Type?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId)) continue;
            var part = mainPart.GetPartById(relationshipId);
            var paragraphs = part switch
            {
                HeaderPart { Header: { } root } => root.Elements<W.Paragraph>().ToArray(),
                FooterPart { Footer: { } root } => root.Elements<W.Paragraph>().ToArray(),
                _ => [],
            };
            var parsed = new List<DocumentHeaderFooter>(paragraphs.Length);
            for (var index = 0; index < paragraphs.Length; index++)
            {
                if (!TryReadParagraph(paragraphs[index], out var text, out var styleId, out var fieldInstruction))
                {
                    parsed.Clear();
                    diagnostics.Add(CodecProtocol.Warning(
                        "unsupported_document_header_footer_preserved",
                        $"Preserved complex {(header ? "header" : "footer")} part {part.Uri} without exposing an editable model.",
                        part.Uri.OriginalString.TrimStart('/')));
                    break;
                }
                var artifact = new DocumentHeaderFooter
                {
                    Id = $"document/{(header ? "header" : "footer")}/{sectionIndex}/{ReferenceToken(type)}/{index + 1}",
                    Name = $"{(header ? "Header" : "Footer")} {index + 1}",
                    StyleId = styleId,
                    Text = text,
                    Reference = FromNativeReference(type),
                    SectionIndex = sectionIndex,
                    RelationshipId = relationshipId,
                    PartPath = part.Uri.OriginalString.TrimStart('/'),
                    VariantActive = type == W.HeaderFooterValues.First
                        ? properties.GetFirstChild<W.TitlePage>() is not null
                        : type == W.HeaderFooterValues.Even ? document.EvenAndOddHeaders : true,
                    FieldInstruction = fieldInstruction,
                };
                parsed.Add(artifact);
            }
            if (header) document.Headers.Add(parsed);
            else document.Footers.Add(parsed);
        }
    }

    private static bool TryReadParagraph(W.Paragraph paragraph, out string text, out string styleId, out string fieldInstruction)
    {
        text = string.Empty;
        styleId = string.Empty;
        fieldInstruction = string.Empty;
        if (paragraph.ParagraphProperties?.ChildElements.Any(child => child is not W.ParagraphStyleId) == true) return false;
        styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? string.Empty;
        var fields = paragraph.Elements<W.SimpleField>().ToArray();
        if (fields.Length == 1 && paragraph.ChildElements.All(child => child is W.ParagraphProperties or W.SimpleField))
        {
            var field = fields[0];
            if (field.Elements<W.Run>().Count() != 1 || field.Descendants<W.RunProperties>().Any()) return false;
            fieldInstruction = field.Instruction?.Value?.Trim() ?? string.Empty;
            text = string.Concat(field.Descendants<W.Text>().Select(item => item.Text));
            try { DocxFieldCodec.Validate(new DocumentField { Instruction = fieldInstruction, Display = text }); }
            catch (CodecException) { return false; }
            return true;
        }
        if (fields.Length > 0 || paragraph.ChildElements.Any(child => child is not W.ParagraphProperties and not W.Run) ||
            paragraph.Descendants<W.RunProperties>().Any() || paragraph.Descendants<W.Run>().Any(run => run.ChildElements.Any(child => child is not W.Text)))
            return false;
        text = string.Concat(paragraph.Descendants<W.Text>().Select(item => item.Text));
        return true;
    }

    private static W.Paragraph BuildParagraph(DocumentHeaderFooter source)
    {
        var paragraph = new W.Paragraph();
        if (!string.IsNullOrWhiteSpace(source.StyleId))
            paragraph.ParagraphProperties = new W.ParagraphProperties(new W.ParagraphStyleId { Val = source.StyleId });
        var run = new W.Run(Text(source.Text));
        if (string.IsNullOrWhiteSpace(source.FieldInstruction)) paragraph.Append(run);
        else paragraph.Append(new W.SimpleField(run) { Instruction = source.FieldInstruction });
        return paragraph;
    }

    private static IEnumerable<Group> Groups(DocumentArtifact document, uint sectionCount)
    {
        foreach (var (header, items) in new[] { (true, document.Headers), (false, document.Footers) })
            foreach (var group in items.GroupBy(item => (
                         SectionIndex: item.HasSectionIndex ? item.SectionIndex : sectionCount - 1,
                         item.Reference)))
                yield return new Group(header, group.Key.SectionIndex, group.Key.Reference, group.ToArray());
    }

    private static W.HeaderFooterValues ToNativeReference(DocumentHeaderFooterReference value) => value switch
    {
        DocumentHeaderFooterReference.First => W.HeaderFooterValues.First,
        DocumentHeaderFooterReference.Even => W.HeaderFooterValues.Even,
        _ => W.HeaderFooterValues.Default,
    };

    private static DocumentHeaderFooterReference FromNativeReference(W.HeaderFooterValues? value) =>
        value == W.HeaderFooterValues.First ? DocumentHeaderFooterReference.First :
        value == W.HeaderFooterValues.Even ? DocumentHeaderFooterReference.Even :
        DocumentHeaderFooterReference.Default;

    private static string ReferenceToken(W.HeaderFooterValues? value) =>
        value == W.HeaderFooterValues.First ? "first" :
        value == W.HeaderFooterValues.Even ? "even" : "default";

    private static W.Text Text(string value) => new(value)
    {
        Space = value.Length != value.Trim().Length ? SpaceProcessingModeValues.Preserve : null,
    };

    private static bool SequenceEqual<T>(IEnumerable<T> left, IEnumerable<T> right) where T : class =>
        left.SequenceEqual(right, EqualityComparer<T>.Default);
}
