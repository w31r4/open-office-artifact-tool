using System.Security.Cryptography;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using V = DocumentFormat.OpenXml.Vml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Owns one deliberately narrow text-watermark profile. A watermark is a
// complete header paragraph containing one VML WordArt text-path shape whose
// ID says WaterMark. Treating the paragraph as the unit of topology lets a
// source-bound removal delete exactly that object while a residual hash proves
// that every other header node stayed untouched.
internal static class DocxWatermarkCodec
{
    private sealed record HeaderReference(
        uint SectionIndex,
        DocumentHeaderFooterReference Reference,
        string RelationshipId,
        HeaderPart Part);

    internal static void Read(
        MainDocumentPart mainPart,
        W.Body body,
        DocumentArtifact document,
        ICollection<Diagnostic> diagnostics)
    {
        var references = HeaderReferences(mainPart, body).ToArray();
        var partUseCounts = references
            .GroupBy(reference => PartPath(reference.Part), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);

        foreach (var reference in references)
        {
            var root = reference.Part.Header;
            if (root is null) continue;
            var candidates = root.Elements<W.Paragraph>()
                .Select((paragraph, index) => (Paragraph: paragraph, Index: index))
                .Where(item => TryReadCanonical(item.Paragraph, out _, out _))
                .ToArray();
            if (candidates.Length == 0) continue;

            var partPath = PartPath(reference.Part);
            if (partUseCounts[partPath] != 1)
            {
                diagnostics.Add(CodecProtocol.Warning(
                    "shared_document_watermark_preserved",
                    $"Preserved text-watermark content in shared Header part {partPath} without exposing mutable semantics.",
                    partPath));
                continue;
            }
            if (candidates.Length != 1)
            {
                diagnostics.Add(CodecProtocol.Warning(
                    "multiple_document_watermarks_preserved",
                    $"Preserved {candidates.Length} text-watermark paragraphs in Header part {partPath}; the bounded model permits one watermark per section/reference scope.",
                    partPath));
                continue;
            }

            var candidate = candidates[0];
            if (!TryReadCanonical(candidate.Paragraph, out var text, out var shapeId)) continue;
            var watermark = new DocumentWatermark
            {
                Id = $"document/watermark/{reference.SectionIndex}/{ReferenceToken(reference.Reference)}",
                Text = text,
                Reference = reference.Reference,
                SectionIndex = reference.SectionIndex,
                Source = new DocumentWatermarkSourceBinding
                {
                    RelationshipId = reference.RelationshipId,
                    PartPath = partPath,
                    ParagraphIndex = checked((uint)candidate.Index),
                    ShapeId = shapeId,
                    ElementSha256 = HashElement(candidate.Paragraph),
                    ResidualSha256 = ResidualHash(candidate.Paragraph),
                    PartResidualSha256 = PartResidualHash(root),
                    Editable = true,
                },
            };
            watermark.Source.SemanticSha256 = SemanticHash(watermark);
            document.Watermarks.Add(watermark);
        }
    }

    internal static W.Paragraph Build(DocumentWatermark watermark, uint ordinal)
    {
        var shapeType = new V.Shapetype
        {
            Id = "_x0000_t136",
            CoordinateSize = "21600,21600",
            OptionalNumber = 136,
            Adjustment = "10800",
            EdgePath = "m@7,l@8,m@5,21600l@6,21600e",
        };
        shapeType.Append(new V.Path { AllowTextPath = true });

        var shape = new V.Shape
        {
            Id = $"OpenChestnutWaterMarkObject{ordinal}",
            Type = "#_x0000_t136",
            Style = "position:absolute;margin-left:0;margin-top:0;width:468pt;height:117pt;rotation:315;z-index:-251654144;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin;mso-wrap-edited:f",
            FillColor = "#C0C0C0",
            Stroked = false,
            AllowInCell = false,
        };
        shape.Append(
            new V.Fill { Opacity = ".18" },
            new V.TextPath
            {
                Style = "font-family:&quot;Calibri&quot;;font-size:1pt",
                String = watermark.Text,
            },
            new V.Path { AllowTextPath = true });

        return new W.Paragraph(
            new W.Run(
                new W.RunProperties(new W.NoProof()),
                new W.Picture(shapeType, shape)));
    }

    internal static bool IsCanonicalParagraph(W.Paragraph paragraph) =>
        TryReadCanonical(paragraph, out _, out _);

    internal static void ApplySource(
        DocxPartContext context,
        W.Body body,
        DocumentArtifact requested)
    {
        var source = new DocumentArtifact();
        var diagnostics = new List<Diagnostic>();
        Read(context.Owner, body, source, diagnostics);
        var requestedById = requested.Watermarks.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var sourceById = source.Watermarks.ToDictionary(item => item.Id, StringComparer.Ordinal);
        foreach (var item in requested.Watermarks)
        {
            if (!sourceById.ContainsKey(item.Id))
                throw new CodecException(
                    "document_watermark_topology_changed",
                    "Source-preserving DOCX export cannot add a watermark; only recognized existing watermarks may be edited or removed.");
        }

        foreach (var original in source.Watermarks)
        {
            requestedById.TryGetValue(original.Id, out var next);
            ApplyOne(context, original, next);
        }
    }

    internal static void Validate(DocumentArtifact document)
    {
        var sectionCount = DocxHeaderFooterCodec.SectionCount(document);
        var ids = new HashSet<string>(StringComparer.Ordinal);
        var scopes = new HashSet<(uint SectionIndex, DocumentHeaderFooterReference Reference)>();
        foreach (var watermark in document.Watermarks)
        {
            if (string.IsNullOrWhiteSpace(watermark.Id) || watermark.Id.Length > 512 || watermark.Id.Any(char.IsControl) || !ids.Add(watermark.Id))
                throw new CodecException("invalid_document_watermark", "Document watermarks require unique IDs of 1 through 512 characters without controls.");
            if (string.IsNullOrWhiteSpace(watermark.Text) || watermark.Text.Length > 256 || !IsXmlSafe(watermark.Text))
                throw new CodecException("invalid_document_watermark", $"Document watermark {watermark.Id} text must contain 1 through 256 XML-safe characters and cannot be blank.");
            if (watermark.Reference is DocumentHeaderFooterReference.Unspecified)
                throw new CodecException("invalid_document_watermark", $"Document watermark {watermark.Id} requires a default, first, or even header reference.");
            if (!watermark.HasSectionIndex || watermark.SectionIndex >= sectionCount)
                throw new CodecException("invalid_document_watermark", $"Document watermark {watermark.Id} section index is outside 0 through {sectionCount - 1}.");
            if (!scopes.Add((watermark.SectionIndex, watermark.Reference)))
                throw new CodecException("invalid_document_watermark", $"Document section {watermark.SectionIndex} has more than one {watermark.Reference} watermark.");
            if (watermark.Source is { } source &&
                (string.IsNullOrWhiteSpace(source.RelationshipId) ||
                 string.IsNullOrWhiteSpace(source.PartPath) ||
                 string.IsNullOrWhiteSpace(source.ShapeId) ||
                 string.IsNullOrWhiteSpace(source.ElementSha256) ||
                 string.IsNullOrWhiteSpace(source.SemanticSha256) ||
                 string.IsNullOrWhiteSpace(source.ResidualSha256) ||
                 string.IsNullOrWhiteSpace(source.PartResidualSha256) ||
                 !source.Editable))
                throw new CodecException("invalid_document_watermark", $"Document watermark {watermark.Id} has an incomplete or read-only source binding.");
        }
    }

    private static void ApplyOne(DocxPartContext context, DocumentWatermark original, DocumentWatermark? requested)
    {
        var binding = original.Source ?? throw new CodecException(
            "document_watermark_source_binding_mismatch",
            $"Document watermark {original.Id} is missing its source binding.");
        if (!binding.Editable)
            throw new CodecException(
                "unsupported_document_watermark_edit",
                $"Document watermark {original.Id} is source-bound and read-only.",
                binding.PartPath);
        if (requested is not null)
        {
            if (requested.Source is null || !BindingEquals(binding, requested.Source) ||
                requested.Reference != original.Reference || !requested.HasSectionIndex || requested.SectionIndex != original.SectionIndex)
                throw new CodecException(
                    "document_watermark_source_binding_mismatch",
                    $"Document watermark {original.Id} source identity, section, or header reference changed.",
                    binding.PartPath);
            if (!SemanticHash(original).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "document_watermark_source_binding_mismatch",
                    $"Document watermark {original.Id} source semantics do not match its binding.",
                    binding.PartPath);
        }

        var part = context.Owner.HeaderParts.SingleOrDefault(candidate =>
            PartPath(candidate).Equals(binding.PartPath, StringComparison.OrdinalIgnoreCase)) ??
            throw new CodecException(
                "document_watermark_source_binding_mismatch",
                $"Document watermark {original.Id} Header part is missing.",
                binding.PartPath);
        var root = part.Header ?? throw new CodecException(
            "document_watermark_source_binding_mismatch",
            $"Document watermark {original.Id} Header part has no root.",
            binding.PartPath);
        var paragraphs = root.Elements<W.Paragraph>().ToArray();
        if (binding.ParagraphIndex >= paragraphs.Length)
            throw new CodecException(
                "document_watermark_source_binding_mismatch",
                $"Document watermark {original.Id} paragraph locator is outside its Header part.",
                binding.PartPath);
        var paragraph = paragraphs[binding.ParagraphIndex];
        if (!HashElement(paragraph).Equals(binding.ElementSha256, StringComparison.OrdinalIgnoreCase) ||
            !ResidualHash(paragraph).Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase) ||
            !PartResidualHash(root).Equals(binding.PartResidualSha256, StringComparison.OrdinalIgnoreCase) ||
            !TryReadCanonical(paragraph, out var sourceText, out var shapeId) ||
            !shapeId.Equals(binding.ShapeId, StringComparison.Ordinal) ||
            !sourceText.Equals(original.Text, StringComparison.Ordinal))
            throw new CodecException(
                "document_watermark_source_binding_mismatch",
                $"Document watermark {original.Id} no longer matches its exact source paragraph.",
                binding.PartPath);

        if (requested is not null && requested.Text.Equals(original.Text, StringComparison.Ordinal)) return;

        var residualBefore = PartResidualHash(root);
        if (requested is null) paragraph.Remove();
        else paragraph.Descendants<V.TextPath>().Single().String = requested.Text;
        if (!PartResidualHash(root).Equals(residualBefore, StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "document_watermark_residual_not_preserved",
                $"Document watermark {original.Id} edit changed unrelated Header content.",
                binding.PartPath);
        if (requested is not null &&
            (!TryReadCanonical(paragraph, out var verifiedText, out var verifiedShapeId) ||
             !verifiedText.Equals(requested.Text, StringComparison.Ordinal) ||
             !verifiedShapeId.Equals(binding.ShapeId, StringComparison.Ordinal)))
            throw new CodecException(
                "document_watermark_semantics_not_applied",
                $"Document watermark {original.Id} text edit did not produce the requested native semantics.",
                binding.PartPath);
        part.Header.Save();
        context.MarkPartMutated(part);
    }

    private static bool TryReadCanonical(W.Paragraph paragraph, out string text, out string shapeId)
    {
        text = string.Empty;
        shapeId = string.Empty;
        var runs = paragraph.Elements<W.Run>().ToArray();
        if (runs.Length != 1 || paragraph.ChildElements.Any(child => child is not W.ParagraphProperties and not W.Run)) return false;
        var run = runs[0];
        if (run.ChildElements.Any(child => child is not W.RunProperties and not W.Picture)) return false;
        var pictures = run.Elements<W.Picture>().ToArray();
        if (pictures.Length != 1) return false;
        var picture = pictures[0];
        if (picture.ChildElements.Any(child => child is not V.Shapetype and not V.Shape)) return false;
        var shapes = picture.Elements<V.Shape>().ToArray();
        if (shapes.Length != 1) return false;
        var shape = shapes[0];
        shapeId = shape.Id?.Value ?? string.Empty;
        if (!shapeId.Contains("watermark", StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(shape.Type?.Value, "#_x0000_t136", StringComparison.OrdinalIgnoreCase) ||
            shape.Descendants<V.ImageData>().Any()) return false;
        var paths = shape.Elements<V.TextPath>().ToArray();
        if (paths.Length != 1) return false;
        text = paths[0].String?.Value ?? string.Empty;
        return !string.IsNullOrWhiteSpace(text) && text.Length <= 256;
    }

    private static IEnumerable<HeaderReference> HeaderReferences(MainDocumentPart mainPart, W.Body body)
    {
        var sections = BoundarySections(body).ToArray();
        for (var sectionIndex = 0; sectionIndex < sections.Length; sectionIndex++)
        {
            foreach (var reference in sections[sectionIndex].Elements<W.HeaderReference>())
            {
                var relationshipId = reference.Id?.Value;
                if (string.IsNullOrWhiteSpace(relationshipId)) continue;
                if (mainPart.GetPartById(relationshipId) is not HeaderPart part) continue;
                yield return new HeaderReference(
                    checked((uint)sectionIndex),
                    FromNativeReference(reference.Type?.Value),
                    relationshipId,
                    part);
            }
        }
    }

    private static IEnumerable<W.SectionProperties> BoundarySections(W.Body body)
    {
        foreach (var paragraph in body.Elements<W.Paragraph>())
            if (paragraph.ParagraphProperties?.SectionProperties is { } boundary) yield return boundary;
        if (body.Elements<W.SectionProperties>().LastOrDefault() is { } final) yield return final;
        else yield return new W.SectionProperties();
    }

    private static DocumentHeaderFooterReference FromNativeReference(W.HeaderFooterValues? value) =>
        value == W.HeaderFooterValues.First ? DocumentHeaderFooterReference.First :
        value == W.HeaderFooterValues.Even ? DocumentHeaderFooterReference.Even :
        DocumentHeaderFooterReference.Default;

    private static string ReferenceToken(DocumentHeaderFooterReference value) => value switch
    {
        DocumentHeaderFooterReference.First => "first",
        DocumentHeaderFooterReference.Even => "even",
        _ => "default",
    };

    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');

    private static string ResidualHash(W.Paragraph paragraph)
    {
        var clone = (W.Paragraph)paragraph.CloneNode(true);
        var textPath = clone.Descendants<V.TextPath>().SingleOrDefault();
        if (textPath is not null) textPath.String = string.Empty;
        return HashElement(clone);
    }

    private static string PartResidualHash(W.Header header)
    {
        var clone = (W.Header)header.CloneNode(true);
        foreach (var paragraph in clone.Elements<W.Paragraph>().Where(IsCanonicalParagraph).ToArray())
            paragraph.Remove();
        return HashElement(clone);
    }

    private static string SemanticHash(DocumentWatermark watermark)
    {
        var semantic = watermark.Clone();
        semantic.Id = string.Empty;
        semantic.Source = null;
        return Hash(semantic.ToByteArray());
    }

    private static bool BindingEquals(DocumentWatermarkSourceBinding left, DocumentWatermarkSourceBinding right) =>
        left.RelationshipId.Equals(right.RelationshipId, StringComparison.Ordinal) &&
        left.PartPath.Equals(right.PartPath, StringComparison.OrdinalIgnoreCase) &&
        left.ParagraphIndex == right.ParagraphIndex &&
        left.ShapeId.Equals(right.ShapeId, StringComparison.Ordinal) &&
        left.ElementSha256.Equals(right.ElementSha256, StringComparison.OrdinalIgnoreCase) &&
        left.SemanticSha256.Equals(right.SemanticSha256, StringComparison.OrdinalIgnoreCase) &&
        left.ResidualSha256.Equals(right.ResidualSha256, StringComparison.OrdinalIgnoreCase) &&
        left.PartResidualSha256.Equals(right.PartResidualSha256, StringComparison.OrdinalIgnoreCase) &&
        left.Editable == right.Editable;

    private static string HashElement(OpenXmlElement element) => Hash(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

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
    }
}
