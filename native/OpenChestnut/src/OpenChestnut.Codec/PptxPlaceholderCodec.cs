using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenChestnut.Codec;

// Owns direct p:ph identity, its local a:txBody, and a bounded direct a:xfrm.
// Fully recognized transform slots may add/remove the frame; less exact but
// readable transforms retain coordinate/rotation/flip editing only. Fills,
// shape style, and inherited/effective formatting remain in the source element.
internal static class PptxPlaceholderCodec
{
    private const int MaxRotationAngle60000 = 360 * 60_000;
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
                    DirectFramePresenceEditable = SupportsDirectFramePresenceEditing(shape),
                },
            };
            placeholder.Source.Editable = placeholder.Source.Editable || placeholder.Source.DirectFramePresenceEditable;
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

    // Source-free authoring has no inherited native geometry to resolve. The
    // caller therefore supplies an explicit direct frame and this builder owns
    // only the canonical text-placeholder p:sp profile; richer placeholder
    // graphs remain source-bound and use Apply instead.
    internal static P.Shape Build(
        PresentationPlaceholder source,
        uint nativeId,
        PptxPartContext partContext)
    {
        Validate(source);
        if (source.DirectFrame is null)
            throw new CodecException("invalid_presentation_placeholder", $"Source-free presentation placeholder {source.Id} requires a direct frame.");

        var nativePlaceholder = new P.PlaceholderShape { Index = source.Index };
        nativePlaceholder.SetAttribute(new OpenXmlAttribute("type", string.Empty, source.Type));
        var output = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = nativeId, Name = source.Name },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties(nativePlaceholder)),
            new P.ShapeProperties(
                new A.Transform2D(new A.Offset(), new A.Extents()),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
                new A.NoFill()),
            PptxTextCodec.Build(TextShape(source.TextBody), partContext));
        ApplyDirectFrame(output, source.DirectFrame);
        return output;
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
        if (source.DirectFrame is not null && requested.DirectFrame is not null)
        {
            if (!SupportsDirectFrame(sourceShape))
                throw new CodecException("unsupported_presentation_edit", $"Presentation placeholder {requested.Id} cannot replace an unrecognized direct frame in this codec slice.");
            ValidateDirectFrame(requested.DirectFrame, requested.Id);
            ApplyDirectFrame(sourceShape, requested.DirectFrame);
            return;
        }
        if (source.Source?.DirectFramePresenceEditable != true || !SupportsDirectFramePresenceEditing(sourceShape))
            throw new CodecException("unsupported_presentation_edit", $"Presentation placeholder {requested.Id} cannot add or remove an unrecognized direct frame in this codec slice.");
        if (requested.DirectFrame is null)
        {
            sourceShape.ShapeProperties!.Transform2D!.Remove();
            return;
        }
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
            if (SupportsDirectFramePresenceEditing(shape))
                shape.ShapeProperties!.Transform2D?.Remove();
            else if (SupportsDirectFrame(shape))
                ScrubDirectFrame(shape);
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

    internal static PresentationPlaceholderIdentity? ReadIdentity(P.Shape shape)
    {
        var native = NativePlaceholder(shape);
        if (native is null) return null;
        return new PresentationPlaceholderIdentity
        {
            Type = NativeType(native),
            Index = native.Index?.Value ?? 0U,
            InheritsGeometry = shape.ShapeProperties?.Transform2D is null,
        };
    }

    internal static PresentationPlaceholderFrame? ReadDirectFrame(P.Shape shape)
    {
        if (!SupportsDirectFrame(shape)) return null;
        var transform = shape.ShapeProperties!.Transform2D!;
        var frame = new PresentationPlaceholderFrame
        {
            LeftEmu = transform.Offset!.X!.Value,
            TopEmu = transform.Offset.Y!.Value,
            WidthEmu = transform.Extents!.Cx!.Value,
            HeightEmu = transform.Extents.Cy!.Value,
        };
        if (transform.Rotation?.Value is { } rotation) frame.RotationAngle60000 = rotation;
        if (transform.HorizontalFlip?.Value is { } flipHorizontal) frame.FlipHorizontal = flipHorizontal;
        if (transform.VerticalFlip?.Value is { } flipVertical) frame.FlipVertical = flipVertical;
        return frame;
    }

    private static bool SupportsDirectFrame(P.Shape shape)
    {
        var transform = shape.ShapeProperties?.Transform2D;
        if (transform is null || transform.ChildElements.Count != 2 ||
            transform.ChildElements[0] is not A.Offset offset ||
            transform.ChildElements[1] is not A.Extents extents ||
            offset.X is null || offset.Y is null || extents.Cx is null || extents.Cy is null)
            return false;
        if (offset.X.Value < 0 || offset.Y.Value < 0 || extents.Cx.Value <= 0 || extents.Cy.Value <= 0) return false;
        return transform.Rotation?.Value is not { } rotation || Math.Abs((long)rotation) <= MaxRotationAngle60000;
    }

    private static bool SupportsDirectFramePresenceEditing(P.Shape shape)
    {
        var properties = shape.ShapeProperties;
        if (properties is null) return false;
        var transforms = properties.Elements<A.Transform2D>().ToArray();
        if (transforms.Length == 0) return true;
        if (transforms.Length != 1 || !SupportsDirectFrame(shape)) return false;
        var transform = transforms[0];
        return HasOnlyAttributes(transform, "rot", "flipH", "flipV") &&
            HasOnlyAttributes(transform.Offset!, "x", "y") &&
            HasOnlyAttributes(transform.Extents!, "cx", "cy");
    }

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute =>
            string.IsNullOrEmpty(attribute.NamespaceUri) && allowed.Contains(attribute.LocalName));
    }

    internal static void ValidateDirectFrame(PresentationPlaceholderFrame? frame, string placeholderId)
    {
        if (frame is null) return;
        if (frame.LeftEmu < 0 || frame.TopEmu < 0 || frame.WidthEmu <= 0 || frame.HeightEmu <= 0)
            throw new CodecException("invalid_presentation_frame", $"Presentation placeholder {placeholderId} has an invalid direct frame.");
        if (frame.HasRotationAngle60000 && Math.Abs((long)frame.RotationAngle60000) > MaxRotationAngle60000)
            throw new CodecException("invalid_presentation_transform", $"Presentation placeholder {placeholderId} rotation must be between -360 and 360 degrees.");
    }

    private static bool FrameEquals(PresentationPlaceholderFrame? left, PresentationPlaceholderFrame? right) =>
        left is null ? right is null : right is not null && left.Equals(right);

    private static void ApplyDirectFrame(P.Shape shape, PresentationPlaceholderFrame frame)
    {
        var properties = shape.ShapeProperties!;
        var transform = properties.Transform2D;
        if (transform is null)
        {
            transform = new A.Transform2D(new A.Offset(), new A.Extents());
            properties.Transform2D = transform;
        }
        transform.Offset!.X = frame.LeftEmu;
        transform.Offset.Y = frame.TopEmu;
        transform.Extents!.Cx = frame.WidthEmu;
        transform.Extents.Cy = frame.HeightEmu;
        transform.Rotation = frame.HasRotationAngle60000 ? frame.RotationAngle60000 : null;
        transform.HorizontalFlip = frame.HasFlipHorizontal ? frame.FlipHorizontal : null;
        transform.VerticalFlip = frame.HasFlipVertical ? frame.FlipVertical : null;
    }

    private static void ScrubDirectFrame(P.Shape shape)
    {
        var transform = shape.ShapeProperties!.Transform2D!;
        transform.Offset!.X = 0L;
        transform.Offset.Y = 0L;
        transform.Extents!.Cx = 1L;
        transform.Extents.Cy = 1L;
        transform.Rotation = null;
        transform.HorizontalFlip = null;
        transform.VerticalFlip = null;
    }

    private static P.PlaceholderShape? NativePlaceholder(P.Shape shape) =>
        shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>();

    private static string NativeType(P.PlaceholderShape source)
    {
        // `p:ph/@type` is a typed attribute; Open XML SDK 3.x rejects the
        // generic empty-namespace lookup for generated elements. The typed
        // property covers the native enum values and preserves the default
        // object placeholder semantics when it is omitted.
        var token = source.Type?.InnerText;
        return string.IsNullOrWhiteSpace(token) ? "obj" : token;
    }

    private static OpenXmlElement[] ShapeElements(P.ShapeTree shapeTree) =>
        shapeTree.ChildElements.Where(child => child is not P.NonVisualGroupShapeProperties and not P.GroupShapeProperties).ToArray();

    private static string HashElement(OpenXmlElement element) => Hash(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
