using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenChestnut.Codec;

// Owns the deliberately small top-level p:pic contract. Everything outside
// the embedded blip, rectangular stretch frame, alt text, and direct visual
// transform remains opaque and must survive the residual hash unchanged.
internal static class PptxPictureCodec
{
    private const int MaxTextLength = 1_024;

    internal static bool TryRead(P.Picture source, PptxPartContext context, out PresentationImage image)
    {
        image = new PresentationImage();
        if (!TryParts(source, out var nonVisual, out var blip, out var transform)) return false;
        var relationshipId = blip.Embed?.Value ?? string.Empty;
        if (relationshipId.Length == 0) return false;
        try
        {
            var asset = context.ReadEmbeddedPicture(relationshipId);
            var offset = transform.Offset!;
            var extents = transform.Extents!;
            image = new PresentationImage
            {
                AssetId = asset.Id,
                AltText = nonVisual.Description?.Value ?? string.Empty,
                LeftEmu = offset.X?.Value ?? 0,
                TopEmu = offset.Y?.Value ?? 0,
                WidthEmu = extents.Cx?.Value ?? 0,
                HeightEmu = extents.Cy?.Value ?? 0,
            };
            var visual = ReadTransform(transform);
            if (visual is not null) image.Transform = visual;
            return image.LeftEmu >= 0 && image.TopEmu >= 0 && image.WidthEmu > 0 && image.HeightEmu > 0 &&
                   (nonVisual.Name?.Value?.Length ?? 0) <= MaxTextLength && image.AltText.Length <= MaxTextLength;
        }
        catch (CodecException)
        {
            image = new PresentationImage();
            return false;
        }
    }

    internal static void Validate(PresentationImage? image, string elementId, PptxAssetCatalog assets)
    {
        if (image is null)
            throw Invalid(elementId, "payload is missing");
        if (string.IsNullOrWhiteSpace(image.AssetId) || image.AssetId.Length > 512)
            throw Invalid(elementId, "asset ID must contain 1 through 512 characters");
        _ = assets.Get(image.AssetId);
        if (image.AltText.Length > MaxTextLength)
            throw Invalid(elementId, $"alternative text exceeds {MaxTextLength} characters");
        if (image.LeftEmu < 0 || image.TopEmu < 0 || image.WidthEmu <= 0 || image.HeightEmu <= 0)
            throw Invalid(elementId, "frame must use non-negative coordinates and positive extents");
        ValidateTransform(image.Transform, elementId);
    }

    internal static P.Picture Build(PresentationElement source, uint nativeId, PptxPartContext context)
    {
        var image = source.Image;
        var transform = new A.Transform2D(
            new A.Offset { X = image.LeftEmu, Y = image.TopEmu },
            new A.Extents { Cx = image.WidthEmu, Cy = image.HeightEmu });
        ApplyTransform(transform, image.Transform);
        return new P.Picture(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties
                {
                    Id = nativeId,
                    Name = source.Name,
                    Description = image.AltText,
                },
                new P.NonVisualPictureDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(
                new A.Blip { Embed = context.AddEmbeddedPicture(image.AssetId) },
                new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(
                transform,
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
    }

    internal static void Apply(P.Picture source, PresentationElement requested, PptxPartContext context)
    {
        if (!TryParts(source, out var nonVisual, out var blip, out var transform))
            throw new CodecException("unsupported_presentation_edit", $"Presentation image {requested.Id} no longer matches the editable picture profile.");
        var current = context.ReadEmbeddedPicture(blip.Embed?.Value ?? string.Empty);
        var replacement = context.Assets?.Get(requested.Image.AssetId) ??
            throw new CodecException("invalid_presentation_asset", $"Presentation image {requested.Id} requires an asset catalog.");
        if (!current.Id.Equals(replacement.Id, StringComparison.Ordinal))
        {
            if (!current.ContentType.Equals(replacement.ContentType, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("unsupported_presentation_image", $"Presentation image {requested.Id} replacement must retain content type {current.ContentType}.");
            blip.Embed = context.AddEmbeddedPicture(replacement.Id);
        }
        nonVisual.Name = requested.Name;
        nonVisual.Description = requested.Image.AltText;
        transform.Offset!.X = requested.Image.LeftEmu;
        transform.Offset.Y = requested.Image.TopEmu;
        transform.Extents!.Cx = requested.Image.WidthEmu;
        transform.Extents.Cy = requested.Image.HeightEmu;
        ApplyTransform(transform, requested.Image.Transform);
    }

    internal static void ScrubModeledContent(P.Picture source)
    {
        if (source.NonVisualPictureProperties?.NonVisualDrawingProperties is { } nonVisual)
        {
            nonVisual.Name = string.Empty;
            nonVisual.Description = string.Empty;
        }
        if (source.BlipFill?.GetFirstChild<A.Blip>() is { } blip) blip.Embed = string.Empty;
        if (source.ShapeProperties?.Transform2D is { } transform)
        {
            if (transform.Offset is { } offset) { offset.X = 0L; offset.Y = 0L; }
            if (transform.Extents is { } extents) { extents.Cx = 1L; extents.Cy = 1L; }
            transform.Rotation = null;
            transform.HorizontalFlip = null;
            transform.VerticalFlip = null;
        }
    }

    private static bool TryParts(
        P.Picture source,
        out P.NonVisualDrawingProperties nonVisual,
        out A.Blip blip,
        out A.Transform2D transform)
    {
        nonVisual = null!;
        blip = null!;
        transform = null!;
        var nonVisualContainer = source.NonVisualPictureProperties;
        var fill = source.BlipFill;
        var properties = source.ShapeProperties;
        if (nonVisualContainer?.NonVisualDrawingProperties is not { } nv ||
            nonVisualContainer.NonVisualPictureDrawingProperties is null ||
            nonVisualContainer.ApplicationNonVisualDrawingProperties is null ||
            fill is null || properties is null) return false;
        if (source.ChildElements.Count != 3 ||
            nonVisualContainer.ChildElements.Count != 3 ||
            fill.ChildElements.Count != 2 ||
            fill.Elements<A.Blip>().SingleOrDefault() is not { } embedded ||
            fill.Elements<A.Stretch>().SingleOrDefault() is not { } stretch ||
            stretch.ChildElements.Count != 1 || stretch.GetFirstChild<A.FillRectangle>() is not { } fillRect ||
            fillRect.HasAttributes || fillRect.HasChildren ||
            embedded.Link is not null || embedded.Embed is null || embedded.ChildElements.Count != 0 || embedded.CompressionState is not null ||
            embedded.GetAttributes().Count != 1 || fill.HasAttributes || stretch.HasAttributes ||
            properties.ChildElements.Count != 2 ||
            properties.Elements<A.Transform2D>().SingleOrDefault() is not { } xfrm ||
            properties.Elements<A.PresetGeometry>().SingleOrDefault() is not { } geometry ||
            geometry.Preset?.Value != A.ShapeTypeValues.Rectangle || geometry.HasAttributes && geometry.GetAttributes().Count != 1 ||
            geometry.ChildElements.Count != 1 || geometry.GetFirstChild<A.AdjustValueList>() is not { } adjustments ||
            adjustments.HasAttributes || adjustments.HasChildren ||
            !TransformSupported(xfrm)) return false;
        nonVisual = nv;
        blip = embedded;
        transform = xfrm;
        return true;
    }

    private static bool TransformSupported(A.Transform2D? transform)
    {
        if (transform is null || transform.ChildElements.Count != 2 ||
            transform.Elements<A.Offset>().SingleOrDefault() is not { } offset ||
            transform.Elements<A.Extents>().SingleOrDefault() is not { } extents ||
            offset.X is null || offset.Y is null || extents.Cx is null || extents.Cy is null ||
            offset.HasChildren || extents.HasChildren || offset.GetAttributes().Count != 2 || extents.GetAttributes().Count != 2)
            return false;
        var known = new HashSet<string>(StringComparer.Ordinal) { "rot", "flipH", "flipV" };
        if (transform.GetAttributes().Any(attribute => !known.Contains(attribute.LocalName))) return false;
        var rotation = transform.Rotation?.Value;
        return rotation is null || Math.Abs((long)rotation.Value) <= 21_600_000L;
    }

    private static PresentationImageTransform? ReadTransform(A.Transform2D source)
    {
        var result = new PresentationImageTransform();
        if (source.Rotation is not null) result.RotationAngle60000 = source.Rotation.Value;
        if (source.HorizontalFlip is not null) result.FlipHorizontal = source.HorizontalFlip.Value;
        if (source.VerticalFlip is not null) result.FlipVertical = source.VerticalFlip.Value;
        return result.CalculateSize() == 0 ? null : result;
    }

    private static void ValidateTransform(PresentationImageTransform? transform, string elementId)
    {
        if (transform is null) return;
        if (!transform.HasRotationAngle60000 && !transform.HasFlipHorizontal && !transform.HasFlipVertical)
            throw Invalid(elementId, "transform must define rotation or a flip");
        if (transform.HasRotationAngle60000 && Math.Abs((long)transform.RotationAngle60000) > 21_600_000L)
            throw Invalid(elementId, "rotation must be between -360 and 360 degrees");
    }

    private static void ApplyTransform(A.Transform2D target, PresentationImageTransform? source)
    {
        target.Rotation = source?.HasRotationAngle60000 == true ? source.RotationAngle60000 : null;
        target.HorizontalFlip = source?.HasFlipHorizontal == true ? source.FlipHorizontal : null;
        target.VerticalFlip = source?.HasFlipVertical == true ? source.FlipVertical : null;
    }

    private static CodecException Invalid(string elementId, string message) =>
        new("invalid_presentation_image", $"Presentation image {elementId} {message}.");
}
