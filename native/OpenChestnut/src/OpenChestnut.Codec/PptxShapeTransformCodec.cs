using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OpenChestnut.Codec;

// Owns only the optional rot/flipH/flipV attributes of one ordinary p:sp
// a:xfrm. Offset/extents remain PresentationShape frame fields. The exact
// recognition gate keeps attributed coordinates, duplicate children, and
// unknown transform extensions outside the editable semantic projection.
internal static class PptxShapeTransformCodec
{
    private const int MaxRotationAngle60000 = 360 * 60_000;

    internal static bool Supports(A.Transform2D? transform)
    {
        if (transform is null || transform.ChildElements.Count != 2 ||
            transform.ChildElements[0] is not A.Offset offset ||
            transform.ChildElements[1] is not A.Extents extents ||
            offset.X is null || offset.Y is null || extents.Cx is null || extents.Cy is null)
            return false;
        if (offset.X.Value < 0 || offset.Y.Value < 0 || extents.Cx.Value <= 0 || extents.Cy.Value <= 0)
            return false;
        if (!HasOnlyAttributes(transform, "rot", "flipH", "flipV") ||
            !HasOnlyAttributes(offset, "x", "y") ||
            !HasOnlyAttributes(extents, "cx", "cy"))
            return false;
        return transform.Rotation?.Value is not { } rotation ||
            Math.Abs((long)rotation) <= MaxRotationAngle60000;
    }

    internal static PresentationShapeTransform? Read(A.Transform2D transform)
    {
        var semantic = new PresentationShapeTransform();
        if (transform.Rotation?.Value is { } rotation) semantic.RotationAngle60000 = rotation;
        if (transform.HorizontalFlip?.Value is { } flipHorizontal) semantic.FlipHorizontal = flipHorizontal;
        if (transform.VerticalFlip?.Value is { } flipVertical) semantic.FlipVertical = flipVertical;
        return HasAnyField(semantic) ? semantic : null;
    }

    internal static void Validate(PresentationShapeTransform? transform, string shapeId)
    {
        if (transform is null) return;
        if (!HasAnyField(transform))
            throw new CodecException("invalid_presentation_transform", $"Presentation shape {shapeId} transform must define rotation or a flip.");
        if (transform.HasRotationAngle60000 && Math.Abs((long)transform.RotationAngle60000) > MaxRotationAngle60000)
            throw new CodecException("invalid_presentation_transform", $"Presentation shape {shapeId} rotation must be between -360 and 360 degrees.");
    }

    internal static void Apply(A.Transform2D target, PresentationShapeTransform? requested)
    {
        target.Rotation = requested?.HasRotationAngle60000 == true ? requested.RotationAngle60000 : null;
        target.HorizontalFlip = requested?.HasFlipHorizontal == true ? requested.FlipHorizontal : null;
        target.VerticalFlip = requested?.HasFlipVertical == true ? requested.FlipVertical : null;
    }

    internal static void Scrub(A.Transform2D target)
    {
        target.Rotation = null;
        target.HorizontalFlip = null;
        target.VerticalFlip = null;
    }

    private static bool HasAnyField(PresentationShapeTransform source) =>
        source.HasRotationAngle60000 || source.HasFlipHorizontal || source.HasFlipVertical;

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute =>
            string.IsNullOrEmpty(attribute.NamespaceUri) && allowed.Contains(attribute.LocalName));
    }
}
