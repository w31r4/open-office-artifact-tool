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
            FileName = $"document-image-{digest[..16]}.{(contentType == "image/png" ? "png" : "jpg")}",
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
        _ => throw Invalid($"Unsupported document image content type {contentType}."),
    };

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
            _ => false,
        };
        if (!valid) throw Invalid($"{label} bytes do not match a supported PNG or JPEG content type.");
    }

    private static string Normalize(string value) => value.Equals("image/jpg", StringComparison.OrdinalIgnoreCase)
        ? "image/jpeg"
        : value.ToLowerInvariant();
    private static string Hash(byte[] data) => Convert.ToHexString(SHA256.HashData(data)).ToLowerInvariant();
    private static CodecException Invalid(string message) => new("invalid_document_image_asset", message);
}

internal static class DocxImageCodec
{
    private const string PictureUri = "http://schemas.openxmlformats.org/drawingml/2006/picture";

    internal static bool TryRead(W.Paragraph paragraph, DocxPartContext context, out DocumentImage image)
    {
        image = new DocumentImage();
        if (!DocxFormattingCodec.IsSupportedParagraphProperties(paragraph.ParagraphProperties) ||
            paragraph.ChildElements.Any(child => child is not W.ParagraphProperties and not W.Run)) return false;
        var runs = paragraph.Elements<W.Run>().Take(2).ToArray();
        if (runs.Length != 1) return false;
        var run = runs[0];
        if (run.ChildElements.Any(child => child is not W.RunProperties and not W.Drawing)) return false;
        var drawings = run.Elements<W.Drawing>().Take(2).ToArray();
        if (drawings.Length != 1) return false;
        var inlines = drawings[0].Elements<WP.Inline>().Take(2).ToArray();
        if (inlines.Length != 1) return false;
        var inline = inlines[0];
        var blips = inline.Descendants<A.Blip>().Take(2).ToArray();
        var pictures = inline.Descendants<Pic.Picture>().Take(2).ToArray();
        if (inline.GetFirstChild<WP.Extent>() is not { } extent ||
            inline.GetFirstChild<WP.DocProperties>() is not { } docProperties ||
            blips.Length != 1 ||
            blips[0].Embed?.Value is not { Length: > 0 } relationshipId ||
            pictures.Length != 1) return false;
        if (extent.Cx?.Value is not > 0 || extent.Cy?.Value is not > 0 ||
            extent.Cx.Value > 95_250_000_000L || extent.Cy.Value > 95_250_000_000L) return false;
        try
        {
            if (context.Owner.GetPartById(relationshipId) is not ImagePart part) return false;
            var asset = context.Images?.Import(part) ?? throw new CodecException("invalid_document_image_asset", "Document image import has no asset catalog.");
            image = new DocumentImage
            {
                AssetId = asset.Id,
                AltText = docProperties.Description?.Value ?? string.Empty,
                WidthEmu = extent.Cx.Value,
                HeightEmu = extent.Cy.Value,
            };
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
        var graphic = new A.Graphic(new A.GraphicData(picture) { Uri = PictureUri });
        var inline = new WP.Inline(
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
        var paragraph = new W.Paragraph();
        var properties = DocxFormattingCodec.BuildParagraphProperties(source.StyleId, null);
        if (properties is not null) paragraph.ParagraphProperties = properties;
        paragraph.Append(new W.Run(new W.Drawing(inline)));
        return paragraph;
    }

    internal static void Apply(W.Paragraph paragraph, DocumentBlock requested, DocxPartContext context)
    {
        Validate(requested.Image, requested.Id, context.Images);
        if (!TryParts(paragraph, out var extent, out var docProperties, out var nonVisual, out var blip))
            throw new CodecException("unsupported_document_edit", $"Document image {requested.Id} no longer matches the editable inline-picture profile.", "word/document.xml");
        var relationshipId = blip.Embed?.Value ?? string.Empty;
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
        extent.Cx = requested.Image.WidthEmu;
        extent.Cy = requested.Image.HeightEmu;
        docProperties.Description = requested.Image.AltText;
        nonVisual.Description = requested.Image.AltText;
        var transform = paragraph.Descendants<A.Transform2D>().SingleOrDefault();
        if (transform?.Extents is { } nativeExtent)
        {
            nativeExtent.Cx = requested.Image.WidthEmu;
            nativeExtent.Cy = requested.Image.HeightEmu;
        }
    }

    internal static void Validate(DocumentImage? image, string blockId, DocxImageAssetCatalog? assets)
    {
        if (image is null) throw Invalid(blockId, "payload is missing");
        if (string.IsNullOrWhiteSpace(image.AssetId) || image.AssetId.Length > 512) throw Invalid(blockId, "asset ID must contain 1 through 512 characters");
        if (assets is null) throw Invalid(blockId, "requires an asset catalog");
        _ = assets.Get(image.AssetId);
        if (image.AltText.Length > 32_767 || image.AltText.Any(char.IsControl)) throw Invalid(blockId, "alternative text must contain at most 32767 characters without controls");
        if (image.WidthEmu is <= 0 or > 95_250_000_000L || image.HeightEmu is <= 0 or > 95_250_000_000L)
            throw Invalid(blockId, "width and height must be positive bounded EMU values");
    }

    private static bool TryParts(
        W.Paragraph paragraph,
        out WP.Extent extent,
        out WP.DocProperties docProperties,
        out Pic.NonVisualDrawingProperties nonVisual,
        out A.Blip blip)
    {
        extent = null!;
        docProperties = null!;
        nonVisual = null!;
        blip = null!;
        var inline = paragraph.Descendants<WP.Inline>().SingleOrDefault();
        if (inline?.GetFirstChild<WP.Extent>() is not { } nativeExtent ||
            inline.GetFirstChild<WP.DocProperties>() is not { } nativeProperties ||
            inline.Descendants<Pic.NonVisualDrawingProperties>().SingleOrDefault() is not { } nativeNonVisual ||
            inline.Descendants<A.Blip>().SingleOrDefault() is not { Embed.Value: { Length: > 0 } } nativeBlip) return false;
        extent = nativeExtent;
        docProperties = nativeProperties;
        nonVisual = nativeNonVisual;
        blip = nativeBlip;
        return true;
    }

    private static CodecException Invalid(string blockId, string message) =>
        new("invalid_document_image", $"Document image {blockId} {message}.");
}
