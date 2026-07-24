using System.Security.Cryptography;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
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
    private sealed record HeaderGroup(uint SectionIndex, DocumentHeaderFooterReference Reference, IReadOnlyList<DocumentHeaderFooter> Blocks, IReadOnlyList<DocumentWatermark> Watermarks);

    internal static void Read(
        MainDocumentPart mainPart,
        W.Body body,
        DocumentArtifact document,
        ICollection<Diagnostic> diagnostics)
    {
        var sections = BoundarySections(body).ToArray();
        var partUseCounts = HeaderFooterPartUseCounts(mainPart, sections);
        for (var sectionIndex = 0; sectionIndex < sections.Length; sectionIndex++)
        {
            var properties = sections[sectionIndex];
            if (properties.GetFirstChild<W.TitlePage>() is not null)
                document.SectionSettings.Add(new DocumentSectionSettings
                {
                    SectionIndex = checked((uint)sectionIndex),
                    DifferentFirstPage = true,
                });
            ReadReferences(mainPart, properties, checked((uint)sectionIndex), header: true, document, diagnostics, partUseCounts);
            ReadReferences(mainPart, properties, checked((uint)sectionIndex), header: false, document, diagnostics, partUseCounts);
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

        foreach (var group in HeaderGroups(document, sectionCount))
        {
            var activeFirst = group.Reference == DocumentHeaderFooterReference.First &&
                (group.Watermarks.Count > 0 || group.Blocks.Any(block => !block.HasVariantActive || block.VariantActive));
            if (activeFirst && document.SectionSettings.All(item => item.SectionIndex != group.SectionIndex))
                plan.SetDifferentFirstPage(group.SectionIndex, true);

            var part = mainPart.AddNewPart<HeaderPart>();
            part.Header = new W.Header(
                group.Blocks.Select(BuildParagraph)
                    .Concat(group.Watermarks.Select((watermark, index) => DocxWatermarkCodec.Build(watermark, checked((uint)index + 1U)))));
            part.Header.Save();
            plan.AddReference(group.SectionIndex, new W.HeaderReference
            {
                Type = ToNativeReference(group.Reference),
                Id = mainPart.GetIdOfPart(part),
            });
        }

        foreach (var group in Groups(document, sectionCount).Where(group => !group.Header))
        {
            var activeFirst = group.Reference == DocumentHeaderFooterReference.First &&
                group.Blocks.Any(block => !block.HasVariantActive || block.VariantActive);
            if (activeFirst && document.SectionSettings.All(item => item.SectionIndex != group.SectionIndex))
                plan.SetDifferentFirstPage(group.SectionIndex, true);
            var part = mainPart.AddNewPart<FooterPart>();
            part.Footer = new W.Footer(group.Blocks.Select(BuildParagraph));
            part.Footer.Save();
            plan.AddReference(group.SectionIndex, new W.FooterReference
            {
                Type = ToNativeReference(group.Reference),
                Id = mainPart.GetIdOfPart(part),
            });
        }

        return plan;
    }

    internal static void ApplySource(
        DocxPartContext context,
        W.Body body,
        DocumentArtifact requested,
        DocumentArtifact? sourceSnapshot = null)
    {
        var source = sourceSnapshot ?? new DocumentArtifact();
        if (sourceSnapshot is null)
        {
            var ignored = new List<Diagnostic>();
            DocxSettingsCodec.Read(context.Owner, source);
            Read(context.Owner, body, source, ignored);
        }
        if (!SequenceEqual(source.SectionSettings, requested.SectionSettings))
            throw new CodecException(
                "unsupported_document_header_footer_edit",
                "Source-preserving DOCX export requires the imported header/footer section activation settings to remain unchanged.",
                "word/document.xml");
        ApplyCollection(context, header: true, source.Headers, requested.Headers);
        ApplyCollection(context, header: false, source.Footers, requested.Footers);
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
            if (item.Segments.Count > 0)
                ValidateSegments(item, kind);
            else if (!string.IsNullOrWhiteSpace(item.FieldInstruction))
                DocxFieldCodec.Validate(new DocumentField { Instruction = item.FieldInstruction, Display = item.Text });
            if (item.Source is { } source &&
                (string.IsNullOrWhiteSpace(source.RelationshipId) ||
                 string.IsNullOrWhiteSpace(source.PartPath) ||
                 string.IsNullOrWhiteSpace(source.ElementSha256) ||
                 string.IsNullOrWhiteSpace(source.SemanticSha256) ||
                 string.IsNullOrWhiteSpace(source.ResidualSha256) ||
                 string.IsNullOrWhiteSpace(source.PartResidualSha256)))
                throw new CodecException("invalid_document_header_footer", $"Document {kind} {item.Id} has an incomplete source binding.");
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

    internal static Dictionary<string, int> HeaderFooterPartUseCounts(
        MainDocumentPart mainPart,
        IReadOnlyList<W.SectionProperties> sections)
    {
        var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var active = new Dictionary<(bool Header, W.HeaderFooterValues Type), OpenXmlPart>();
        var duplicateParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var properties in sections)
        {
            var explicitReferences = new Dictionary<(bool Header, W.HeaderFooterValues Type), List<OpenXmlPart>>();
            foreach (var reference in properties.Elements<W.HeaderReference>())
            {
                if (reference.Type?.Value is not { } type || string.IsNullOrWhiteSpace(reference.Id?.Value)) continue;
                var part = mainPart.GetPartById(reference.Id!.Value!);
                if (part is not HeaderPart) continue;
                AddExplicitReference(explicitReferences, (true, type), part);
            }
            foreach (var reference in properties.Elements<W.FooterReference>())
            {
                if (reference.Type?.Value is not { } type || string.IsNullOrWhiteSpace(reference.Id?.Value)) continue;
                var part = mainPart.GetPartById(reference.Id!.Value!);
                if (part is not FooterPart) continue;
                AddExplicitReference(explicitReferences, (false, type), part);
            }

            foreach (var header in new[] { true, false })
            foreach (var type in new[] { W.HeaderFooterValues.Default, W.HeaderFooterValues.First, W.HeaderFooterValues.Even })
            {
                var key = (header, type);
                if (explicitReferences.TryGetValue(key, out var parts))
                {
                    if (parts.Count == 1) active[key] = parts[0];
                    else
                    {
                        active.Remove(key);
                        foreach (var part in parts) duplicateParts.Add(PartPath(part));
                    }
                }
                if (!active.TryGetValue(key, out var activePart)) continue;
                var partPath = PartPath(activePart);
                counts[partPath] = counts.GetValueOrDefault(partPath) + 1;
            }
        }
        foreach (var partPath in duplicateParts)
            counts[partPath] = Math.Max(2, counts.GetValueOrDefault(partPath));
        return counts;
    }

    private static void AddExplicitReference(
        IDictionary<(bool Header, W.HeaderFooterValues Type), List<OpenXmlPart>> references,
        (bool Header, W.HeaderFooterValues Type) key,
        OpenXmlPart part)
    {
        if (!references.TryGetValue(key, out var parts)) references[key] = parts = [];
        parts.Add(part);
    }

    private static void ReadReferences(
        MainDocumentPart mainPart,
        W.SectionProperties properties,
        uint sectionIndex,
        bool header,
        DocumentArtifact document,
        ICollection<Diagnostic> diagnostics,
        IReadOnlyDictionary<string, int> partUseCounts)
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
            var root = part switch
            {
                HeaderPart { Header: { } headerRoot } => (OpenXmlCompositeElement)headerRoot,
                FooterPart { Footer: { } footerRoot } => footerRoot,
                _ => null,
            };
            if (root is null) continue;
            var paragraphs = root.Elements<W.Paragraph>().ToArray();
            var partPath = PartPath(part);
            var parsed = new List<DocumentHeaderFooter>(paragraphs.Length);
            for (var index = 0; index < paragraphs.Length; index++)
            {
                if (header && DocxWatermarkCodec.IsCanonicalParagraph(paragraphs[index])) continue;
                if (!TryReadParagraph(paragraphs[index], out var text, out var styleId, out var fieldInstruction, out var segments))
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
                    PartPath = partPath,
                    VariantActive = type == W.HeaderFooterValues.First
                        ? properties.GetFirstChild<W.TitlePage>() is not null
                        : type == W.HeaderFooterValues.Even ? document.EvenAndOddHeaders : true,
                    FieldInstruction = fieldInstruction,
                };
                artifact.Segments.Add(segments);
                artifact.Source = new DocumentHeaderFooterSourceBinding
                {
                    RelationshipId = relationshipId,
                    PartPath = partPath,
                    ParagraphIndex = checked((uint)index),
                    ElementSha256 = HashElement(paragraphs[index]),
                    ResidualSha256 = ResidualHash(paragraphs[index]),
                    PartResidualSha256 = PartResidualHash(root, checked((uint)index)),
                    Editable = partUseCounts.GetValueOrDefault(partPath) == 1 && CanEditText(paragraphs[index]),
                };
                artifact.Source.SemanticSha256 = SemanticHash(artifact);
                parsed.Add(artifact);
            }
            if (header) document.Headers.Add(parsed);
            else document.Footers.Add(parsed);
        }
    }

    private static void ApplyCollection(
        DocxPartContext context,
        bool header,
        IList<DocumentHeaderFooter> source,
        IList<DocumentHeaderFooter> requested)
    {
        if (source.Count != requested.Count)
            throw new CodecException(
                "document_header_footer_topology_changed",
                $"Source-preserving DOCX export requires the original {(header ? "header" : "footer")} topology.",
                "word/document.xml");

        var changes = new List<(DocumentHeaderFooter Original, DocumentHeaderFooter Requested)>();
        for (var index = 0; index < source.Count; index++)
        {
            var original = source[index];
            var next = requested[index];
            var binding = original.Source;
            if (binding is null || !IdentityEquals(original, next) ||
                !SemanticHash(original).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "document_header_footer_source_binding_mismatch",
                    $"Document {(header ? "header" : "footer")} {original.Id} no longer matches its source binding.",
                    binding?.PartPath ?? original.PartPath);
            if (original.Text.Equals(next.Text, StringComparison.Ordinal) && SegmentsEqual(original.Segments, next.Segments)) continue;
            if (!binding.Editable)
                throw new CodecException(
                    "unsupported_document_header_footer_edit",
                    $"Document {(header ? "header" : "footer")} {original.Id} is source-bound and cannot replace its text in this codec profile.",
                    binding.PartPath);
            changes.Add((original, next));
        }

        foreach (var group in changes.GroupBy(change => change.Original.Source!.PartPath, StringComparer.OrdinalIgnoreCase))
            if (group.Count() > 1)
                throw new CodecException(
                    "document_header_footer_multiple_edits",
                    $"Source-preserving DOCX export permits at most one text edit per imported {(header ? "Header" : "Footer")} part.",
                    group.Key);

        foreach (var (original, next) in changes)
            ApplyOne(context, header, original, next);
    }

    private static void ApplyOne(
        DocxPartContext context,
        bool header,
        DocumentHeaderFooter original,
        DocumentHeaderFooter requested)
    {
        var binding = original.Source ?? throw new CodecException(
            "document_header_footer_source_binding_mismatch",
            $"Document {(header ? "header" : "footer")} {original.Id} is missing its source binding.",
            original.PartPath);
        if (!binding.Editable)
            throw new CodecException(
                "unsupported_document_header_footer_edit",
                $"Document {(header ? "header" : "footer")} {original.Id} is source-bound and cannot replace its text in this codec profile.",
                binding.PartPath);
        if (!binding.RelationshipId.Equals(original.RelationshipId, StringComparison.Ordinal) ||
            !binding.PartPath.Equals(original.PartPath, StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "document_header_footer_source_binding_mismatch",
                $"Document {(header ? "header" : "footer")} {original.Id} source relationship identity changed.",
                binding.PartPath);

        var part = context.Owner.GetPartById(binding.RelationshipId);
        OpenXmlCompositeElement? root = null;
        if (header && part is HeaderPart headerPart) root = headerPart.Header;
        if (!header && part is FooterPart footerPart) root = footerPart.Footer;
        if (root is null || !PartPath(part).Equals(binding.PartPath, StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "document_header_footer_source_binding_mismatch",
                $"Document {(header ? "header" : "footer")} {original.Id} part is missing or has the wrong kind.",
                binding.PartPath);

        var paragraphs = root.Elements<W.Paragraph>().ToArray();
        W.Paragraph paragraph;
        uint paragraphIndex;
        if (binding.ParagraphIndex < paragraphs.Length &&
            MatchesSourceParagraph(paragraphs[binding.ParagraphIndex], original, binding))
        {
            paragraph = paragraphs[binding.ParagraphIndex];
            paragraphIndex = binding.ParagraphIndex;
        }
        else
        {
            // A prior recognized watermark removal can shift the paragraph
            // index in this HeaderPart. The original binding remains the
            // authority, so only one exact hash/semantic match may re-anchor.
            var matches = paragraphs.Select((candidate, index) => (Paragraph: candidate, Index: index))
                .Where(candidate => MatchesSourceParagraph(candidate.Paragraph, original, binding))
                .Take(2)
                .ToArray();
            if (matches.Length != 1)
                throw new CodecException(
                    "document_header_footer_source_binding_mismatch",
                    $"Document {(header ? "header" : "footer")} {original.Id} paragraph locator no longer identifies one exact source paragraph.",
                    binding.PartPath);
            paragraph = matches[0].Paragraph;
            paragraphIndex = checked((uint)matches[0].Index);
        }
        if (!PartResidualHash(root, paragraphIndex).Equals(binding.PartResidualSha256, StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "document_header_footer_source_binding_mismatch",
                $"Document {(header ? "header" : "footer")} {original.Id} no longer matches its exact source paragraph.",
                binding.PartPath);

        var partResidualBefore = PartResidualHash(root, paragraphIndex);
        var text = paragraph.Descendants<W.Text>().Single();
        text.Text = requested.Text;
        text.Space = requested.Text.Length != requested.Text.Trim().Length ? SpaceProcessingModeValues.Preserve : null;
        if (!PartResidualHash(root, paragraphIndex).Equals(partResidualBefore, StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "document_header_footer_residual_not_preserved",
                $"Document {(header ? "header" : "footer")} {original.Id} text edit changed unrelated source content.",
                binding.PartPath);
        if (!TryReadParagraph(paragraph, out var verifiedText, out var verifiedStyleId, out var verifiedFieldInstruction, out var verifiedSegments) ||
            !CanEditText(paragraph) ||
            !verifiedText.Equals(requested.Text, StringComparison.Ordinal) ||
            !verifiedStyleId.Equals(original.StyleId, StringComparison.Ordinal) ||
            !verifiedFieldInstruction.Equals(original.FieldInstruction, StringComparison.Ordinal) ||
            !SegmentsEqual(verifiedSegments, original.Segments))
            throw new CodecException(
                "document_header_footer_semantics_not_applied",
                $"Document {(header ? "header" : "footer")} {original.Id} text edit did not produce the requested native semantics.",
                binding.PartPath);

        if (part is HeaderPart savedHeader) savedHeader.Header!.Save();
        else ((FooterPart)part).Footer!.Save();
        context.MarkPartMutated(part);
    }

    private static bool IdentityEquals(DocumentHeaderFooter left, DocumentHeaderFooter right) =>
        left.Id.Equals(right.Id, StringComparison.Ordinal) &&
        left.Name.Equals(right.Name, StringComparison.Ordinal) &&
        left.StyleId.Equals(right.StyleId, StringComparison.Ordinal) &&
        left.Reference == right.Reference &&
        left.HasSectionIndex == right.HasSectionIndex &&
        (!left.HasSectionIndex || left.SectionIndex == right.SectionIndex) &&
        left.RelationshipId.Equals(right.RelationshipId, StringComparison.Ordinal) &&
        left.PartPath.Equals(right.PartPath, StringComparison.OrdinalIgnoreCase) &&
        left.HasVariantActive == right.HasVariantActive &&
        (!left.HasVariantActive || left.VariantActive == right.VariantActive) &&
        left.FieldInstruction.Equals(right.FieldInstruction, StringComparison.Ordinal) &&
        SegmentsEqual(left.Segments, right.Segments) &&
        BindingEquals(left.Source, right.Source);

    private static bool BindingEquals(DocumentHeaderFooterSourceBinding? left, DocumentHeaderFooterSourceBinding? right) =>
        left is not null && right is not null &&
        left.RelationshipId.Equals(right.RelationshipId, StringComparison.Ordinal) &&
        left.PartPath.Equals(right.PartPath, StringComparison.OrdinalIgnoreCase) &&
        left.ParagraphIndex == right.ParagraphIndex &&
        left.ElementSha256.Equals(right.ElementSha256, StringComparison.OrdinalIgnoreCase) &&
        left.SemanticSha256.Equals(right.SemanticSha256, StringComparison.OrdinalIgnoreCase) &&
        left.ResidualSha256.Equals(right.ResidualSha256, StringComparison.OrdinalIgnoreCase) &&
        left.PartResidualSha256.Equals(right.PartResidualSha256, StringComparison.OrdinalIgnoreCase) &&
        left.Editable == right.Editable;

    private static bool CanEditText(W.Paragraph paragraph)
    {
        if (!TryReadParagraph(paragraph, out _, out _, out var fieldInstruction, out var segments) ||
            !string.IsNullOrEmpty(fieldInstruction) || segments.Count > 0) return false;
        var runs = paragraph.Elements<W.Run>().ToArray();
        return runs.Length == 1 && runs[0].ChildElements.Count == 1 && runs[0].GetFirstChild<W.Text>() is not null;
    }

    private static bool MatchesSourceParagraph(
        W.Paragraph paragraph,
        DocumentHeaderFooter original,
        DocumentHeaderFooterSourceBinding binding)
    {
        if (!HashElement(paragraph).Equals(binding.ElementSha256, StringComparison.OrdinalIgnoreCase) ||
            !ResidualHash(paragraph).Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase) ||
            !TryReadParagraph(paragraph, out var sourceText, out var styleId, out var fieldInstruction, out var segments) ||
            !CanEditText(paragraph)) return false;
        return sourceText.Equals(original.Text, StringComparison.Ordinal) &&
               styleId.Equals(original.StyleId, StringComparison.Ordinal) &&
               fieldInstruction.Equals(original.FieldInstruction, StringComparison.Ordinal) &&
               SegmentsEqual(segments, original.Segments);
    }

    private static string ResidualHash(W.Paragraph paragraph)
    {
        var clone = (W.Paragraph)paragraph.CloneNode(true);
        foreach (var text in clone.Descendants<W.Text>())
        {
            text.Text = string.Empty;
            text.Space = null;
        }
        return HashElement(clone);
    }

    private static string PartResidualHash(OpenXmlCompositeElement root, uint paragraphIndex)
    {
        var clone = (OpenXmlCompositeElement)root.CloneNode(true);
        var paragraphs = clone.Elements<W.Paragraph>().ToArray();
        for (var index = 0; index < paragraphs.Length; index++)
            if (index == paragraphIndex || DocxWatermarkCodec.IsCanonicalParagraph(paragraphs[index]))
                paragraphs[index].Remove();
        return HashElement(clone);
    }

    private static string SemanticHash(DocumentHeaderFooter source)
    {
        var semantic = source.Clone();
        semantic.Id = string.Empty;
        semantic.Name = string.Empty;
        semantic.Source = null;
        return Hash(semantic.ToByteArray());
    }

    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');

    private static string HashElement(OpenXmlElement element) => Hash(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static bool TryReadParagraph(
        W.Paragraph paragraph,
        out string text,
        out string styleId,
        out string fieldInstruction,
        out IReadOnlyList<DocumentHeaderFooterSegment> segments)
    {
        text = string.Empty;
        styleId = string.Empty;
        fieldInstruction = string.Empty;
        segments = [];
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
        if (fields.Length > 0)
            return TryReadSegments(paragraph, out text, out segments);
        if (paragraph.ChildElements.Any(child => child is not W.ParagraphProperties and not W.Run) ||
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
        if (source.Segments.Count > 0)
        {
            foreach (var segment in source.Segments)
            {
                if (segment.ContentCase == DocumentHeaderFooterSegment.ContentOneofCase.Text)
                    paragraph.Append(new W.Run(Text(segment.Text)));
                else if (segment.ContentCase == DocumentHeaderFooterSegment.ContentOneofCase.Field)
                    paragraph.Append(new W.SimpleField(new W.Run(Text(segment.Field.Display))) { Instruction = segment.Field.Instruction });
            }
            return paragraph;
        }
        var run = new W.Run(Text(source.Text));
        if (string.IsNullOrWhiteSpace(source.FieldInstruction)) paragraph.Append(run);
        else paragraph.Append(new W.SimpleField(run) { Instruction = source.FieldInstruction });
        return paragraph;
    }

    private static void ValidateSegments(DocumentHeaderFooter item, string kind)
    {
        if (!string.IsNullOrEmpty(item.FieldInstruction))
            throw new CodecException("invalid_document_header_footer", $"Document {kind} {item.Id} cannot combine legacy field_instruction with structured segments.");
        if (item.Segments.Count is < 2 or > 32)
            throw new CodecException("invalid_document_header_footer", $"Document {kind} {item.Id} requires 2 through 32 structured page-furniture segments.");

        var text = new StringBuilder();
        var fields = 0;
        foreach (var segment in item.Segments)
        {
            switch (segment.ContentCase)
            {
                case DocumentHeaderFooterSegment.ContentOneofCase.Text:
                    if (string.IsNullOrEmpty(segment.Text) || segment.Text.Length > 1_000_000 ||
                        !IsXmlSafe(segment.Text))
                        throw new CodecException("invalid_document_header_footer", $"Document {kind} {item.Id} has an invalid structured text segment.");
                    text.Append(segment.Text);
                    break;
                case DocumentHeaderFooterSegment.ContentOneofCase.Field:
                    if (segment.Field is null || segment.Field.Complex)
                        throw new CodecException("invalid_document_header_footer", $"Document {kind} {item.Id} structured fields must use one bounded simple-field profile.");
                    DocxFieldCodec.Validate(segment.Field);
                    if (!IsXmlSafe(segment.Field.Display))
                        throw new CodecException("invalid_document_header_footer", $"Document {kind} {item.Id} has an invalid structured field display.");
                    text.Append(segment.Field.Display);
                    fields++;
                    break;
                default:
                    throw new CodecException("invalid_document_header_footer", $"Document {kind} {item.Id} has a structured segment without text or field content.");
            }
            if (text.Length > 1_000_000)
                throw new CodecException("invalid_document_header_footer", $"Document {kind} {item.Id} structured segments exceed 1,000,000 display characters.");
        }
        if (fields == 0 || !text.ToString().Equals(item.Text, StringComparison.Ordinal))
            throw new CodecException("invalid_document_header_footer", $"Document {kind} {item.Id} structured segments must contain a field and exactly match text.");
    }

    private static bool IsXmlSafe(string value)
    {
        try
        {
            XmlConvert.VerifyXmlChars(value);
            return true;
        }
        catch (XmlException)
        {
            return false;
        }
        catch (ArgumentException)
        {
            return false;
        }
    }

    private static bool TryReadSegments(
        W.Paragraph paragraph,
        out string text,
        out IReadOnlyList<DocumentHeaderFooterSegment> segments)
    {
        text = string.Empty;
        segments = [];
        if (paragraph.ChildElements.Any(child => child is not W.ParagraphProperties and not W.Run and not W.SimpleField) ||
            paragraph.Descendants<W.RunProperties>().Any()) return false;

        var parsed = new List<DocumentHeaderFooterSegment>();
        foreach (var child in paragraph.ChildElements)
        {
            switch (child)
            {
                case W.ParagraphProperties:
                    continue;
                case W.Run run when run.ChildElements.Count == 1 && run.GetFirstChild<W.Text>() is { } value && !string.IsNullOrEmpty(value.Text):
                    parsed.Add(new DocumentHeaderFooterSegment { Text = value.Text });
                    break;
                case W.SimpleField field when TryReadSegmentField(field, out var parsedField):
                    parsed.Add(new DocumentHeaderFooterSegment { Field = parsedField });
                    break;
                default:
                    return false;
            }
        }
        if (parsed.Count is < 2 or > 32 || !parsed.Any(segment => segment.ContentCase == DocumentHeaderFooterSegment.ContentOneofCase.Field))
            return false;
        text = string.Concat(parsed.Select(segment => segment.ContentCase == DocumentHeaderFooterSegment.ContentOneofCase.Text
            ? segment.Text
            : segment.Field.Display));
        if (text.Length > 1_000_000) return false;
        segments = parsed;
        return true;
    }

    private static bool TryReadSegmentField(W.SimpleField field, out DocumentField parsed)
    {
        parsed = new DocumentField();
        if (field.Instruction?.Value is not { } instruction || field.ChildElements.Any(child => child is not W.Run)) return false;
        var runs = field.Elements<W.Run>().ToArray();
        if (runs.Length != 1 || runs[0].ChildElements.Count != 1 || runs[0].GetFirstChild<W.Text>() is not { } text) return false;
        parsed.Instruction = instruction;
        parsed.Display = text.Text;
        parsed.Complex = false;
        try
        {
            DocxFieldCodec.Validate(parsed);
            return true;
        }
        catch (CodecException)
        {
            return false;
        }
    }

    private static bool SegmentsEqual(IEnumerable<DocumentHeaderFooterSegment> left, IEnumerable<DocumentHeaderFooterSegment> right) =>
        left.SequenceEqual(right, EqualityComparer<DocumentHeaderFooterSegment>.Default);

    private static IEnumerable<Group> Groups(DocumentArtifact document, uint sectionCount)
    {
        foreach (var (header, items) in new[] { (true, document.Headers), (false, document.Footers) })
            foreach (var group in items.GroupBy(item => (
                         SectionIndex: item.HasSectionIndex ? item.SectionIndex : sectionCount - 1,
                         item.Reference)))
                yield return new Group(header, group.Key.SectionIndex, group.Key.Reference, group.ToArray());
    }

    private static IEnumerable<HeaderGroup> HeaderGroups(DocumentArtifact document, uint sectionCount)
    {
        var headers = document.Headers.GroupBy(item => (
            SectionIndex: item.HasSectionIndex ? item.SectionIndex : sectionCount - 1,
            item.Reference)).ToDictionary(group => group.Key, group => (IReadOnlyList<DocumentHeaderFooter>)group.ToArray());
        var watermarks = document.Watermarks.GroupBy(item => (
            SectionIndex: item.HasSectionIndex ? item.SectionIndex : 0U,
            item.Reference)).ToDictionary(group => group.Key, group => (IReadOnlyList<DocumentWatermark>)group.ToArray());
        foreach (var key in headers.Keys.Concat(watermarks.Keys).Distinct().OrderBy(key => key.SectionIndex).ThenBy(key => (int)key.Reference))
            yield return new HeaderGroup(
                key.SectionIndex,
                key.Reference,
                headers.GetValueOrDefault(key) ?? [],
                watermarks.GetValueOrDefault(key) ?? []);
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
