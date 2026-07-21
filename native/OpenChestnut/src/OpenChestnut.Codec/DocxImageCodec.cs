using System.Globalization;
using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using W = DocumentFormat.OpenXml.Wordprocessing;
using WP = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OpenChestnut.Codec;

internal sealed class DocxImageAssetCatalog
{
    private const int MaxAssets = 1_024;
    private const int MaxAssetBytes = 16 * 1024 * 1024;
    private const string Prefix = "asset/document/image/";
    private readonly Dictionary<string, Asset> _requested = new(StringComparer.Ordinal);
    private readonly Dictionary<string, Asset> _imported = new(StringComparer.Ordinal);
    private readonly ulong _maxTotalBytes;
    private ulong _totalBytes;

    internal DocxImageAssetCatalog(IEnumerable<Asset>? assets, EffectiveCodecLimits limits)
    {
        _maxTotalBytes = Math.Min(limits.MaxUncompressedBytes, (ulong)MaxAssets * MaxAssetBytes);
        foreach (var asset in assets ?? []) AddRequested(asset);
    }

    internal IReadOnlyCollection<Asset> ImportedAssets => _imported.Values;
    internal Asset Get(string id) => _requested.TryGetValue(id, out var asset)
        ? asset
        : throw Invalid($"Document image references missing asset {id}.");

    internal Asset Import(ImagePart part)
    {
        using var input = part.GetStream(FileMode.Open, FileAccess.Read);
        using var output = new MemoryStream();
        input.CopyTo(output);
        var data = output.ToArray();
        var contentType = Normalize(part.ContentType);
        ValidateBytes(contentType, data, $"Document image part {part.Uri}");
        var digest = Hash(data);
        var id = Prefix + digest;
        if (_requested.TryGetValue(id, out var requested)) return requested;
        if (_imported.TryGetValue(id, out var imported)) return imported;
        if (_imported.Count >= MaxAssets)
            throw new CodecException("document_image_asset_budget_exceeded", $"Document exceeds the {MaxAssets}-image asset budget.");
        EnsureBudget(data.LongLength);
        var asset = new Asset
        {
            Id = id,
            FileName = $"document-image-{digest[..16]}.{ExtensionFor(contentType)}",
            ContentType = contentType,
            Data = ByteString.CopyFrom(data),
            Sha256 = digest,
        };
        _imported.Add(id, asset);
        return asset;
    }

    internal static PartTypeInfo PartTypeFor(string contentType) => Normalize(contentType) switch
    {
        "image/png" => ImagePartType.Png,
        "image/jpeg" => ImagePartType.Jpeg,
        "image/gif" => ImagePartType.Gif,
        _ => throw Invalid($"Unsupported document image content type {contentType}."),
    };

    internal static bool IsOrdinaryDocumentImage(Asset asset) =>
        Normalize(asset.ContentType) is "image/png" or "image/jpeg";

    internal static bool IsOrdinaryDocumentImage(ImagePart part) =>
        Normalize(part.ContentType) is "image/png" or "image/jpeg";

    internal static bool SameContentType(ImagePart part, Asset asset) =>
        Normalize(part.ContentType).Equals(Normalize(asset.ContentType), StringComparison.Ordinal);

    private void AddRequested(Asset source)
    {
        if (_requested.Count >= MaxAssets)
            throw new CodecException("document_image_asset_budget_exceeded", $"Document exceeds the {MaxAssets}-image asset budget.");
        var contentType = Normalize(source.ContentType);
        var data = source.Data.ToByteArray();
        ValidateBytes(contentType, data, $"Document asset {source.Id}");
        var digest = Hash(data);
        if (!source.Sha256.Equals(digest, StringComparison.OrdinalIgnoreCase) ||
            !source.Id.Equals(Prefix + digest, StringComparison.Ordinal))
            throw Invalid($"Document asset {source.Id} is not content-addressed by its bytes.");
        if (!_requested.TryAdd(source.Id, source.Clone()))
            throw Invalid($"Document contains duplicate asset ID {source.Id}.");
        EnsureBudget(data.LongLength);
    }

    private void EnsureBudget(long length)
    {
        if (length is <= 0 or > MaxAssetBytes)
            throw new CodecException("document_image_asset_budget_exceeded", $"Document image assets must contain 1 through {MaxAssetBytes} bytes.");
        _totalBytes = checked(_totalBytes + (ulong)length);
        if (_totalBytes > _maxTotalBytes)
            throw new CodecException("document_image_asset_budget_exceeded", $"Document image assets exceed the {_maxTotalBytes}-byte budget.");
    }

    private static void ValidateBytes(string contentType, byte[] data, string label)
    {
        var valid = contentType switch
        {
            "image/png" => data.AsSpan().StartsWith(Convert.FromHexString("89504E470D0A1A0A")),
            "image/jpeg" => data.Length >= 4 && data[0] == 0xff && data[1] == 0xd8 && data[2] == 0xff,
            "image/gif" => data.AsSpan().StartsWith("GIF87a"u8) || data.AsSpan().StartsWith("GIF89a"u8),
            _ => false,
        };
        if (!valid) throw Invalid($"{label} bytes do not match a supported PNG, JPEG, or GIF content type.");
    }

    private static string ExtensionFor(string contentType) => contentType switch
    {
        "image/png" => "png",
        "image/jpeg" => "jpg",
        "image/gif" => "gif",
        _ => "bin",
    };

    private static string Normalize(string value) => value.Equals("image/jpg", StringComparison.OrdinalIgnoreCase)
        ? "image/jpeg"
        : value.ToLowerInvariant();
    private static string Hash(byte[] data) => Convert.ToHexString(SHA256.HashData(data)).ToLowerInvariant();
    private static CodecException Invalid(string message) => new("invalid_document_image_asset", message);
}

internal static class DocxImageCodec
{
    private const string PictureUri = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    private const long MaxCoordinateEmu = 95_250_000L;
    private const uint CanonicalRelativeHeight = 251_658_240U;

    private sealed record NativeImageParts(
        WP.Inline? Inline,
        WP.Anchor? Anchor,
        WP.Extent Extent,
        WP.DocProperties DocProperties,
        Pic.NonVisualDrawingProperties NonVisual,
        A.Blip Blip);

    internal static bool TryRead(W.Paragraph paragraph, DocxPartContext context, out DocumentImage image)
    {
        image = new DocumentImage();
        if (!TryParts(paragraph, out var parts) ||
            parts.Blip.Embed?.Value is not { Length: > 0 } relationshipId ||
            parts.Extent.Cx?.Value is not > 0 || parts.Extent.Cy?.Value is not > 0 ||
            parts.Extent.Cx.Value > 95_250_000_000L || parts.Extent.Cy.Value > 95_250_000_000L) return false;
        try
        {
            if (context.Owner.GetPartById(relationshipId) is not ImagePart part ||
                !DocxImageAssetCatalog.IsOrdinaryDocumentImage(part)) return false;
            var asset = context.Images?.Import(part) ?? throw new CodecException("invalid_document_image_asset", "Document image import has no asset catalog.");
            image = new DocumentImage
            {
                AssetId = asset.Id,
                AltText = parts.DocProperties.Description?.Value ?? string.Empty,
                WidthEmu = parts.Extent.Cx.Value,
                HeightEmu = parts.Extent.Cy.Value,
            };
            if (parts.Anchor is not null)
            {
                if (!TryReadFloating(parts.Anchor, out var floating)) return false;
                image.Floating = floating;
            }
            return image.AltText.Length <= 32_767;
        }
        catch (CodecException)
        {
            image = new DocumentImage();
            return false;
        }
    }

    internal static W.Paragraph Build(DocumentBlock source, DocxPartContext context)
    {
        Validate(source.Image, source.Id, context.Images);
        var asset = context.Images!.Get(source.Image.AssetId);
        var part = context.Owner.AddImagePart(DocxImageAssetCatalog.PartTypeFor(asset.ContentType));
        using (var input = new MemoryStream(asset.Data.ToByteArray(), writable: false)) part.FeedData(input);
        var relationshipId = context.Owner.GetIdOfPart(part);
        var drawingId = context.NextDrawingId();
        var name = string.IsNullOrWhiteSpace(source.Name) ? $"Image {drawingId}" : source.Name;
        var graphic = BuildGraphic(source, relationshipId, drawingId, name);
        OpenXmlElement container = source.Image.Floating is null
            ? BuildInline(source, drawingId, name, graphic)
            : BuildAnchor(source, drawingId, name, graphic);
        var paragraph = new W.Paragraph();
        var properties = DocxFormattingCodec.BuildParagraphProperties(source.StyleId, null);
        if (properties is not null) paragraph.ParagraphProperties = properties;
        paragraph.Append(new W.Run(new W.Drawing(container)));
        return paragraph;
    }

    private static A.Graphic BuildGraphic(DocumentBlock source, string relationshipId, uint drawingId, string name)
    {
        var picture = new Pic.Picture(
            new Pic.NonVisualPictureProperties(
                new Pic.NonVisualDrawingProperties { Id = drawingId, Name = name, Description = source.Image.AltText },
                new Pic.NonVisualPictureDrawingProperties()),
            new Pic.BlipFill(
                new A.Blip { Embed = relationshipId },
                new A.Stretch(new A.FillRectangle())),
            new Pic.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = source.Image.WidthEmu, Cy = source.Image.HeightEmu }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
        return new A.Graphic(new A.GraphicData(picture) { Uri = PictureUri });
    }

    private static WP.Inline BuildInline(DocumentBlock source, uint drawingId, string name, A.Graphic graphic)
    {
        return new WP.Inline(
            new WP.Extent { Cx = source.Image.WidthEmu, Cy = source.Image.HeightEmu },
            new WP.DocProperties { Id = drawingId, Name = name, Description = source.Image.AltText },
            new WP.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
            graphic)
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 0U,
            DistanceFromRight = 0U,
        };
    }

    private static WP.Anchor BuildAnchor(DocumentBlock source, uint drawingId, string name, A.Graphic graphic)
    {
        var floating = source.Image.Floating!;
        var anchor = new WP.Anchor(
            new WP.SimplePosition { X = 0L, Y = 0L },
            BuildHorizontalPosition(floating),
            BuildVerticalPosition(floating),
            new WP.Extent { Cx = source.Image.WidthEmu, Cy = source.Image.HeightEmu },
            BuildWrap(floating),
            new WP.DocProperties { Id = drawingId, Name = name, Description = source.Image.AltText },
            new WP.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
            graphic)
        {
            DistanceFromTop = floating.DistanceTopEmu,
            DistanceFromRight = floating.DistanceRightEmu,
            DistanceFromBottom = floating.DistanceBottomEmu,
            DistanceFromLeft = floating.DistanceLeftEmu,
            SimplePos = false,
            RelativeHeight = CanonicalRelativeHeight,
            BehindDoc = false,
            Locked = false,
            LayoutInCell = true,
            AllowOverlap = false,
        };
        return anchor;
    }

    internal static void Apply(W.Paragraph paragraph, DocumentBlock requested, DocxPartContext context)
    {
        Validate(requested.Image, requested.Id, context.Images);
        if (!TryParts(paragraph, out var parts))
            throw new CodecException("unsupported_document_edit", $"Document image {requested.Id} no longer matches the editable inline/floating-picture profile.", "word/document.xml");
        if ((parts.Anchor is null) != (requested.Image.Floating is null))
            throw new CodecException("unsupported_document_image_edit", $"Document image {requested.Id} cannot change between inline and floating placement after import.", "word/document.xml");
        if (parts.Anchor is not null && !TryReadFloating(parts.Anchor, out _))
            throw new CodecException("unsupported_document_image_edit", $"Document image {requested.Id} floating source graph is outside the editable profile.", "word/document.xml");
        var relationshipId = parts.Blip.Embed?.Value ?? string.Empty;
        if (context.Owner.GetPartById(relationshipId) is not ImagePart part)
            throw new CodecException("document_source_binding_mismatch", $"Document image {requested.Id} relationship does not resolve to an Image part.", "word/document.xml");
        var replacement = context.Images!.Get(requested.Image.AssetId);
        var current = context.Images.Import(part);
        if (!current.Id.Equals(replacement.Id, StringComparison.Ordinal))
        {
            if (!DocxImageAssetCatalog.SameContentType(part, replacement))
                throw new CodecException("unsupported_document_image_edit", $"Document image {requested.Id} replacement must retain source content type {part.ContentType}.");
            using var input = new MemoryStream(replacement.Data.ToByteArray(), writable: false);
            part.FeedData(input);
            context.MarkPartMutated(part);
        }
        var paragraphProperties = paragraph.ParagraphProperties ??= new W.ParagraphProperties();
        paragraphProperties.ParagraphStyleId = string.IsNullOrWhiteSpace(requested.StyleId)
            ? null
            : new W.ParagraphStyleId { Val = requested.StyleId };
        parts.Extent.Cx = requested.Image.WidthEmu;
        parts.Extent.Cy = requested.Image.HeightEmu;
        parts.DocProperties.Description = requested.Image.AltText;
        parts.NonVisual.Description = requested.Image.AltText;
        var transform = paragraph.Descendants<A.Transform2D>().SingleOrDefault();
        if (transform?.Extents is { } nativeExtent)
        {
            nativeExtent.Cx = requested.Image.WidthEmu;
            nativeExtent.Cy = requested.Image.HeightEmu;
        }
        if (parts.Anchor is not null) ApplyFloating(parts.Anchor, requested.Image.Floating!);
    }

    internal static void Validate(DocumentImage? image, string blockId, DocxImageAssetCatalog? assets)
    {
        if (image is null) throw Invalid(blockId, "payload is missing");
        if (string.IsNullOrWhiteSpace(image.AssetId) || image.AssetId.Length > 512) throw Invalid(blockId, "asset ID must contain 1 through 512 characters");
        if (assets is null) throw Invalid(blockId, "requires an asset catalog");
        var asset = assets.Get(image.AssetId);
        if (!DocxImageAssetCatalog.IsOrdinaryDocumentImage(asset))
            throw Invalid(blockId, "ordinary document images support PNG or JPEG assets");
        if (image.AltText.Length > 32_767 || image.AltText.Any(char.IsControl)) throw Invalid(blockId, "alternative text must contain at most 32767 characters without controls");
        if (image.WidthEmu is <= 0 or > 95_250_000_000L || image.HeightEmu is <= 0 or > 95_250_000_000L)
            throw Invalid(blockId, "width and height must be positive bounded EMU values");
        if (image.Floating is not null) ValidateFloating(image.Floating, blockId);
    }

    private static bool TryParts(W.Paragraph paragraph, out NativeImageParts parts)
    {
        parts = null!;
        if (!DocxFormattingCodec.IsSupportedParagraphProperties(paragraph.ParagraphProperties) ||
            paragraph.ChildElements.Any(child => child is not W.ParagraphProperties and not W.Run)) return false;
        var runs = paragraph.Elements<W.Run>().Take(2).ToArray();
        if (runs.Length != 1 || runs[0].ChildElements.Any(child => child is not W.RunProperties and not W.Drawing)) return false;
        var drawings = runs[0].Elements<W.Drawing>().Take(2).ToArray();
        if (drawings.Length != 1 || drawings[0].ChildElements.Any(child => child is not WP.Inline and not WP.Anchor)) return false;
        var inlines = drawings[0].Elements<WP.Inline>().Take(2).ToArray();
        var anchors = drawings[0].Elements<WP.Anchor>().Take(2).ToArray();
        if (inlines.Length + anchors.Length != 1) return false;
        var inline = inlines.SingleOrDefault();
        var anchor = anchors.SingleOrDefault();
        OpenXmlElement container = (OpenXmlElement?)inline ?? anchor!;
        var extents = container.Elements<WP.Extent>().Take(2).ToArray();
        var documentProperties = container.Elements<WP.DocProperties>().Take(2).ToArray();
        var blips = container.Descendants<A.Blip>().Take(2).ToArray();
        var pictures = container.Descendants<Pic.Picture>().Take(2).ToArray();
        var nonVisuals = container.Descendants<Pic.NonVisualDrawingProperties>().Take(2).ToArray();
        if (extents.Length != 1 || documentProperties.Length != 1 ||
            nonVisuals.Length != 1 || blips.Length != 1 || blips[0].Embed?.Value is not { Length: > 0 } || pictures.Length != 1) return false;
        parts = new NativeImageParts(inline, anchor, extents[0], documentProperties[0], nonVisuals[0], blips[0]);
        return true;
    }

    private static bool TryReadFloating(WP.Anchor anchor, out DocumentFloatingImagePlacement floating)
    {
        floating = new DocumentFloatingImagePlacement();
        if (anchor.SimplePos?.Value != false || anchor.BehindDoc?.Value != false || anchor.Locked?.Value != false ||
            anchor.LayoutInCell?.Value != true || anchor.AllowOverlap?.Value != false || anchor.Hidden?.Value == true ||
            anchor.RelativeHeight?.Value is null || anchor.ExtendedAttributes.Any() ||
            anchor.ChildElements.Any(child => child is not WP.SimplePosition and not WP.HorizontalPosition and not WP.VerticalPosition and
                not WP.Extent and not WP.EffectExtent and not WP.WrapSquare and not WP.WrapTopBottom and not WP.DocProperties and
                not WP.NonVisualGraphicFrameDrawingProperties and not A.Graphic)) return false;
        var simplePositions = anchor.Elements<WP.SimplePosition>().Take(2).ToArray();
        var horizontalPositions = anchor.Elements<WP.HorizontalPosition>().Take(2).ToArray();
        var verticalPositions = anchor.Elements<WP.VerticalPosition>().Take(2).ToArray();
        var effectExtents = anchor.Elements<WP.EffectExtent>().Take(2).ToArray();
        if (simplePositions.Length != 1 || simplePositions[0] is not { X.Value: 0, Y.Value: 0 } ||
            simplePositions[0].ExtendedAttributes.Any() || simplePositions[0].ChildElements.Count != 0 ||
            horizontalPositions.Length != 1 || verticalPositions.Length != 1 ||
            horizontalPositions[0] is not { } horizontal || verticalPositions[0] is not { } vertical ||
            horizontal.ExtendedAttributes.Any() || vertical.ExtendedAttributes.Any() ||
            !TryPositionOffset(horizontal, out var horizontalOffset) || !TryPositionOffset(vertical, out var verticalOffset) ||
            effectExtents.Length > 1 || effectExtents.Any(effect => effect.ExtendedAttributes.Any() ||
                effect.ChildElements.Count != 0 || effect.GetAttributes().Any(attribute => attribute.Value != "0"))) return false;
        if (!TryHorizontalReference(horizontal.RelativeFrom?.Value, out var horizontalReference) ||
            !TryVerticalReference(vertical.RelativeFrom?.Value, out var verticalReference) ||
            !TryWrap(anchor, out var wrapMode, out var wrapSide)) return false;
        floating = new DocumentFloatingImagePlacement
        {
            HorizontalRelativeFrom = horizontalReference,
            HorizontalOffsetEmu = horizontalOffset,
            VerticalRelativeFrom = verticalReference,
            VerticalOffsetEmu = verticalOffset,
            WrapMode = wrapMode,
            WrapSide = wrapSide,
            DistanceTopEmu = anchor.DistanceFromTop?.Value ?? 0U,
            DistanceRightEmu = anchor.DistanceFromRight?.Value ?? 0U,
            DistanceBottomEmu = anchor.DistanceFromBottom?.Value ?? 0U,
            DistanceLeftEmu = anchor.DistanceFromLeft?.Value ?? 0U,
        };
        try
        {
            ValidateFloating(floating, "source");
            return true;
        }
        catch (CodecException)
        {
            floating = new DocumentFloatingImagePlacement();
            return false;
        }
    }

    private static bool TryPositionOffset(OpenXmlCompositeElement position, out long value)
    {
        value = 0;
        var offsets = position.Elements<WP.PositionOffset>().Take(2).ToArray();
        return offsets.Length == 1 && position.ChildElements.Count == 1 &&
            !offsets[0].ExtendedAttributes.Any() && offsets[0].ChildElements.Count == 0 &&
            long.TryParse(offsets[0].Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out value) &&
            value >= -MaxCoordinateEmu && value <= MaxCoordinateEmu;
    }

    private static bool TryWrap(WP.Anchor anchor, out DocumentImageWrapMode mode, out DocumentImageWrapSide side)
    {
        mode = DocumentImageWrapMode.Unspecified;
        side = DocumentImageWrapSide.Unspecified;
        var squares = anchor.Elements<WP.WrapSquare>().Take(2).ToArray();
        var topBottom = anchor.Elements<WP.WrapTopBottom>().Take(2).ToArray();
        if (squares.Length + topBottom.Length != 1) return false;
        if (squares.Length == 1)
        {
            if (squares[0].ChildElements.Count != 0 || squares[0].ExtendedAttributes.Any() ||
                !TryWrapSide(squares[0].WrapText?.Value, out side)) return false;
            mode = DocumentImageWrapMode.Square;
            return true;
        }
        if (topBottom[0].ChildElements.Count != 0 || topBottom[0].GetAttributes().Count != 0 || topBottom[0].ExtendedAttributes.Any()) return false;
        mode = DocumentImageWrapMode.TopAndBottom;
        return true;
    }

    private static WP.HorizontalPosition BuildHorizontalPosition(DocumentFloatingImagePlacement value) =>
        new(new WP.PositionOffset(value.HorizontalOffsetEmu.ToString(CultureInfo.InvariantCulture)))
        {
            RelativeFrom = value.HorizontalRelativeFrom switch
            {
                DocumentImageHorizontalRelativeFrom.Margin => WP.HorizontalRelativePositionValues.Margin,
                DocumentImageHorizontalRelativeFrom.Page => WP.HorizontalRelativePositionValues.Page,
                DocumentImageHorizontalRelativeFrom.Column => WP.HorizontalRelativePositionValues.Column,
                _ => throw Invalid("floating", "horizontal reference is unsupported"),
            },
        };

    private static WP.VerticalPosition BuildVerticalPosition(DocumentFloatingImagePlacement value) =>
        new(new WP.PositionOffset(value.VerticalOffsetEmu.ToString(CultureInfo.InvariantCulture)))
        {
            RelativeFrom = value.VerticalRelativeFrom switch
            {
                DocumentImageVerticalRelativeFrom.Margin => WP.VerticalRelativePositionValues.Margin,
                DocumentImageVerticalRelativeFrom.Page => WP.VerticalRelativePositionValues.Page,
                DocumentImageVerticalRelativeFrom.Paragraph => WP.VerticalRelativePositionValues.Paragraph,
                _ => throw Invalid("floating", "vertical reference is unsupported"),
            },
        };

    private static OpenXmlElement BuildWrap(DocumentFloatingImagePlacement value) => value.WrapMode switch
    {
        DocumentImageWrapMode.Square => new WP.WrapSquare { WrapText = value.WrapSide switch
        {
            DocumentImageWrapSide.BothSides => WP.WrapTextValues.BothSides,
            DocumentImageWrapSide.Left => WP.WrapTextValues.Left,
            DocumentImageWrapSide.Right => WP.WrapTextValues.Right,
            DocumentImageWrapSide.Largest => WP.WrapTextValues.Largest,
            _ => throw Invalid("floating", "square wrap side is unsupported"),
        } },
        DocumentImageWrapMode.TopAndBottom => new WP.WrapTopBottom(),
        _ => throw Invalid("floating", "wrap mode is unsupported"),
    };

    private static void ApplyFloating(WP.Anchor anchor, DocumentFloatingImagePlacement value)
    {
        anchor.HorizontalPosition = BuildHorizontalPosition(value);
        anchor.VerticalPosition = BuildVerticalPosition(value);
        anchor.DistanceFromTop = value.DistanceTopEmu;
        anchor.DistanceFromRight = value.DistanceRightEmu;
        anchor.DistanceFromBottom = value.DistanceBottomEmu;
        anchor.DistanceFromLeft = value.DistanceLeftEmu;
        var currentWrap = anchor.ChildElements.Single(child => child is WP.WrapSquare or WP.WrapTopBottom);
        anchor.ReplaceChild(BuildWrap(value), currentWrap);
    }

    private static void ValidateFloating(DocumentFloatingImagePlacement value, string blockId)
    {
        if (value.HorizontalRelativeFrom is not (DocumentImageHorizontalRelativeFrom.Margin or DocumentImageHorizontalRelativeFrom.Page or DocumentImageHorizontalRelativeFrom.Column))
            throw Invalid(blockId, "floating horizontal reference must be margin, page, or column");
        if (value.VerticalRelativeFrom is not (DocumentImageVerticalRelativeFrom.Margin or DocumentImageVerticalRelativeFrom.Page or DocumentImageVerticalRelativeFrom.Paragraph))
            throw Invalid(blockId, "floating vertical reference must be margin, page, or paragraph");
        if (value.HorizontalOffsetEmu < -MaxCoordinateEmu || value.HorizontalOffsetEmu > MaxCoordinateEmu ||
            value.VerticalOffsetEmu < -MaxCoordinateEmu || value.VerticalOffsetEmu > MaxCoordinateEmu)
            throw Invalid(blockId, "floating offsets exceed the bounded coordinate range");
        if (value.DistanceTopEmu > MaxCoordinateEmu || value.DistanceRightEmu > MaxCoordinateEmu ||
            value.DistanceBottomEmu > MaxCoordinateEmu || value.DistanceLeftEmu > MaxCoordinateEmu)
            throw Invalid(blockId, "floating text distances exceed the bounded coordinate range");
        if (value.WrapMode == DocumentImageWrapMode.Square &&
            value.WrapSide is not (DocumentImageWrapSide.BothSides or DocumentImageWrapSide.Left or DocumentImageWrapSide.Right or DocumentImageWrapSide.Largest))
            throw Invalid(blockId, "square wrap requires a supported side");
        if (value.WrapMode == DocumentImageWrapMode.TopAndBottom && value.WrapSide != DocumentImageWrapSide.Unspecified)
            throw Invalid(blockId, "top-and-bottom wrap cannot carry a square-wrap side");
        if (value.WrapMode is not (DocumentImageWrapMode.Square or DocumentImageWrapMode.TopAndBottom))
            throw Invalid(blockId, "floating wrap mode must be square or top-and-bottom");
    }

    private static bool TryHorizontalReference(WP.HorizontalRelativePositionValues? value, out DocumentImageHorizontalRelativeFrom result)
    {
        result = value == WP.HorizontalRelativePositionValues.Margin
            ? DocumentImageHorizontalRelativeFrom.Margin
            : value == WP.HorizontalRelativePositionValues.Page
                ? DocumentImageHorizontalRelativeFrom.Page
                : value == WP.HorizontalRelativePositionValues.Column
                    ? DocumentImageHorizontalRelativeFrom.Column
                    : DocumentImageHorizontalRelativeFrom.Unspecified;
        return result != DocumentImageHorizontalRelativeFrom.Unspecified;
    }

    private static bool TryVerticalReference(WP.VerticalRelativePositionValues? value, out DocumentImageVerticalRelativeFrom result)
    {
        result = value == WP.VerticalRelativePositionValues.Margin
            ? DocumentImageVerticalRelativeFrom.Margin
            : value == WP.VerticalRelativePositionValues.Page
                ? DocumentImageVerticalRelativeFrom.Page
                : value == WP.VerticalRelativePositionValues.Paragraph
                    ? DocumentImageVerticalRelativeFrom.Paragraph
                    : DocumentImageVerticalRelativeFrom.Unspecified;
        return result != DocumentImageVerticalRelativeFrom.Unspecified;
    }

    private static bool TryWrapSide(WP.WrapTextValues? value, out DocumentImageWrapSide result)
    {
        result = value == WP.WrapTextValues.BothSides
            ? DocumentImageWrapSide.BothSides
            : value == WP.WrapTextValues.Left
                ? DocumentImageWrapSide.Left
                : value == WP.WrapTextValues.Right
                    ? DocumentImageWrapSide.Right
                    : value == WP.WrapTextValues.Largest
                        ? DocumentImageWrapSide.Largest
                        : DocumentImageWrapSide.Unspecified;
        return result != DocumentImageWrapSide.Unspecified;
    }

    private static CodecException Invalid(string blockId, string message) =>
        new("invalid_document_image", $"Document image {blockId} {message}.");
}
