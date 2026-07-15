using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenChestnut.Codec;

// Owns direct p:ph identity, its local a:txBody, and the four coordinates of
// an already-present, recognized direct a:xfrm. Rotation, fills, shape style,
// and inherited/effective formatting remain in the source element.
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
            var directFrame = ReadDirectFrame(shape);
            var placeholder = new PresentationPlaceholder
            {
                Id = $"{ownerId}/placeholder/{shapeTreeIndex + 1}",
                Name = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? $"Placeholder {shapeTreeIndex + 1}",
                Type = NativeType(native),
                Index = native.Index?.Value ?? 0U,
                TextBody = textBody,
                DirectFrame = directFrame,
                Source = new PresentationElementSourceBinding
                {
                    ShapeTreeIndex = checked((uint)shapeTreeIndex),
                    ElementSha256 = HashElement(shape),
                    Editable = PptxTextCodec.SupportsEditing(shape.TextBody) || directFrame is not null,
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
        ValidateDirectFrame(source.DirectFrame, source.Id);
        PptxTextCodec.Validate(TextShape(source.TextBody));
    }

    internal static void Apply(
        P.Shape sourceShape,
        PresentationPlaceholder source,
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

        var textChanged = !TextSemanticHash(source.TextBody).Equals(TextSemanticHash(requested.TextBody), StringComparison.OrdinalIgnoreCase);
        if (textChanged)
        {
            if (!PptxTextCodec.SupportsEditing(sourceShape.TextBody))
                throw new CodecException("unsupported_presentation_edit", $"Presentation placeholder {requested.Id} has an unrecognized local text graph.");
            PptxTextCodec.Apply(sourceShape, TextShape(requested.TextBody), partContext);
        }

        if (FrameEquals(source.DirectFrame, requested.DirectFrame)) return;
        if (source.DirectFrame is null || requested.DirectFrame is null || !SupportsDirectFrame(sourceShape))
            throw new CodecException("unsupported_presentation_edit", $"Presentation placeholder {requested.Id} cannot add, remove, or replace an unrecognized direct frame in this codec slice.");
        ValidateDirectFrame(requested.DirectFrame, requested.Id);
        ApplyDirectFrame(sourceShape, requested.DirectFrame);
    }

    internal static void ScrubModeledContent(P.ShapeTree? shapeTree, PptxPartContext partContext)
    {
        if (shapeTree is null) return;
        foreach (var shape in ShapeElements(shapeTree).OfType<P.Shape>())
        {
            if (NativePlaceholder(shape) is null) continue;
            if (PptxTextCodec.SupportsEditing(shape.TextBody)) PptxTextCodec.ScrubModeledContent(shape.TextBody, partContext);
            if (SupportsDirectFrame(shape)) ScrubDirectFrame(shape);
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

    private static string TextSemanticHash(PresentationTextBody? body)
    {
        var shape = TextShape(body);
        PptxTextCodec.NormalizeSemantics(shape);
        return Hash((shape.TextBody ?? new PresentationTextBody()).ToByteArray());
    }

    private static PresentationPlaceholderFrame? ReadDirectFrame(P.Shape shape)
    {
        if (!SupportsDirectFrame(shape)) return null;
        var transform = shape.ShapeProperties!.Transform2D!;
        return new PresentationPlaceholderFrame
        {
            LeftEmu = transform.Offset!.X!.Value,
            TopEmu = transform.Offset.Y!.Value,
            WidthEmu = transform.Extents!.Cx!.Value,
            HeightEmu = transform.Extents.Cy!.Value,
        };
    }

    private static bool SupportsDirectFrame(P.Shape shape)
    {
        var transform = shape.ShapeProperties?.Transform2D;
        if (transform is null || transform.ChildElements.Count != 2 ||
            transform.ChildElements[0] is not A.Offset offset ||
            transform.ChildElements[1] is not A.Extents extents ||
            offset.X is null || offset.Y is null || extents.Cx is null || extents.Cy is null)
            return false;
        return offset.X.Value >= 0 && offset.Y.Value >= 0 && extents.Cx.Value > 0 && extents.Cy.Value > 0;
    }

    private static void ValidateDirectFrame(PresentationPlaceholderFrame? frame, string placeholderId)
    {
        if (frame is null) return;
        if (frame.LeftEmu < 0 || frame.TopEmu < 0 || frame.WidthEmu <= 0 || frame.HeightEmu <= 0)
            throw new CodecException("invalid_presentation_frame", $"Presentation placeholder {placeholderId} has an invalid direct frame.");
    }

    private static bool FrameEquals(PresentationPlaceholderFrame? left, PresentationPlaceholderFrame? right) =>
        left is null ? right is null : right is not null && left.Equals(right);

    private static void ApplyDirectFrame(P.Shape shape, PresentationPlaceholderFrame frame)
    {
        var transform = shape.ShapeProperties!.Transform2D!;
        transform.Offset!.X = frame.LeftEmu;
        transform.Offset.Y = frame.TopEmu;
        transform.Extents!.Cx = frame.WidthEmu;
        transform.Extents.Cy = frame.HeightEmu;
    }

    private static void ScrubDirectFrame(P.Shape shape)
    {
        var transform = shape.ShapeProperties!.Transform2D!;
        transform.Offset!.X = 0L;
        transform.Offset.Y = 0L;
        transform.Extents!.Cx = 1L;
        transform.Extents.Cy = 1L;
    }

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
