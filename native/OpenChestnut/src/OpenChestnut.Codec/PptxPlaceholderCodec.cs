using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenChestnut.Codec;

// Owns direct p:ph identity plus its local a:txBody. Geometry, fills, shape
// style, and inherited/effective formatting remain in the source element.
internal static class PptxPlaceholderCodec
{
    private static readonly HashSet<string> PlaceholderTypes = new(StringComparer.Ordinal)
    {
        "title", "body", "ctrTitle", "subTitle", "dt", "sldNum", "ftr", "hdr",
        "obj", "chart", "tbl", "clipArt", "dgm", "media", "sldImg", "pic",
    };

    internal static IReadOnlyList<PresentationPlaceholder> Read(
        P.ShapeTree shapeTree,
        string ownerId,
        PptxPartContext partContext)
    {
        var placeholders = new List<PresentationPlaceholder>();
        var elements = ShapeElements(shapeTree);
        for (var shapeTreeIndex = 0; shapeTreeIndex < elements.Length; shapeTreeIndex++)
        {
            if (elements[shapeTreeIndex] is not P.Shape shape) continue;
            var native = NativePlaceholder(shape);
            if (native is null) continue;
            var textBody = PptxTextCodec.Read(shape.TextBody, partContext);
            var placeholder = new PresentationPlaceholder
            {
                Id = $"{ownerId}/placeholder/{shapeTreeIndex + 1}",
                Name = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? $"Placeholder {shapeTreeIndex + 1}",
                Type = NativeType(native),
                Index = native.Index?.Value ?? 0U,
                TextBody = textBody,
                Source = new PresentationElementSourceBinding
                {
                    ShapeTreeIndex = checked((uint)shapeTreeIndex),
                    ElementSha256 = HashElement(shape),
                    Editable = PptxTextCodec.SupportsEditing(shape.TextBody),
                },
            };
            placeholder.Source.SemanticSha256 = SemanticHash(placeholder);
            placeholders.Add(placeholder);
        }
        return placeholders;
    }

    internal static string SemanticHash(PresentationPlaceholder? source)
    {
        var semantic = source?.Clone() ?? new PresentationPlaceholder();
        semantic.Id = string.Empty;
        semantic.Source = null;
        var shape = TextShape(semantic.TextBody);
        PptxTextCodec.NormalizeSemantics(shape);
        semantic.TextBody = shape.TextBody;
        return Hash(semantic.ToByteArray());
    }

    internal static void Validate(PresentationPlaceholder source)
    {
        if (string.IsNullOrWhiteSpace(source.Id))
            throw new CodecException("invalid_presentation_placeholder", "Presentation placeholder IDs must be non-empty.");
        if (source.Name.Length > 1_024)
            throw new CodecException("invalid_presentation_placeholder", $"Presentation placeholder {source.Id} name exceeds 1024 characters.");
        if (!PlaceholderTypes.Contains(source.Type))
            throw new CodecException("invalid_presentation_placeholder", $"Presentation placeholder {source.Id} uses unsupported type {source.Type}.");
        PptxTextCodec.Validate(TextShape(source.TextBody));
    }

    internal static void Apply(
        P.Shape sourceShape,
        PresentationPlaceholder requested,
        PptxPartContext partContext)
    {
        var native = NativePlaceholder(sourceShape) ??
            throw new CodecException("presentation_placeholder_binding_mismatch", $"Presentation placeholder {requested.Id} no longer targets a native p:ph shape.");
        var sourceName = sourceShape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ??
            $"Placeholder {(requested.Source?.ShapeTreeIndex ?? 0U) + 1}";
        if (!requested.Name.Equals(sourceName, StringComparison.Ordinal) ||
            !requested.Type.Equals(NativeType(native), StringComparison.Ordinal) ||
            requested.Index != (native.Index?.Value ?? 0U))
            throw new CodecException("unsupported_presentation_edit", $"Presentation placeholder {requested.Id} cannot change name, type, or index in this codec slice.");
        PptxTextCodec.Apply(sourceShape, TextShape(requested.TextBody), partContext);
    }

    internal static void ScrubModeledContent(P.ShapeTree? shapeTree, PptxPartContext partContext)
    {
        if (shapeTree is null) return;
        foreach (var shape in ShapeElements(shapeTree).OfType<P.Shape>())
        {
            if (NativePlaceholder(shape) is null || !PptxTextCodec.SupportsEditing(shape.TextBody)) continue;
            PptxTextCodec.ScrubModeledContent(shape.TextBody, partContext);
        }
    }

    internal static P.Shape? BoundShape(P.ShapeTree shapeTree, PresentationPlaceholder source)
    {
        var elements = ShapeElements(shapeTree);
        if (source.Source is null || source.Source.ShapeTreeIndex > int.MaxValue) return null;
        var index = (int)source.Source.ShapeTreeIndex;
        return index >= 0 && index < elements.Length ? elements[index] as P.Shape : null;
    }

    internal static string ElementHash(P.Shape shape) => HashElement(shape);

    private static PresentationShape TextShape(PresentationTextBody? body) => new()
    {
        TextBody = body?.Clone() ?? new PresentationTextBody(),
        Text = PptxTextCodec.Flatten(body),
    };

    private static P.PlaceholderShape? NativePlaceholder(P.Shape shape) =>
        shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>();

    private static string NativeType(P.PlaceholderShape source)
    {
        var token = source.GetAttribute("type", string.Empty).Value;
        return string.IsNullOrWhiteSpace(token) ? "obj" : token;
    }

    private static OpenXmlElement[] ShapeElements(P.ShapeTree shapeTree) =>
        shapeTree.ChildElements.Where(child => child is not P.NonVisualGroupShapeProperties and not P.GroupShapeProperties).ToArray();

    private static string HashElement(OpenXmlElement element) => Hash(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
