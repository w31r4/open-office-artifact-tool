using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns content-addressed worksheet image identity. Drawing part paths and
// relationship IDs remain package-local source locators.
internal sealed class XlsxImageAssetCatalog
{
    private const int MaxAssets = 1_024;
    private const int MaxAssetBytes = 16 * 1024 * 1024;
    private const string AssetPrefix = "asset/workbook/image/";
    private readonly Dictionary<string, Asset> _requested = new(StringComparer.Ordinal);
    private readonly Dictionary<string, Asset> _imported = new(StringComparer.Ordinal);
    private readonly ulong _maxTotalBytes;
    private ulong _totalBytes;

    internal XlsxImageAssetCatalog(IEnumerable<Asset>? assets, EffectiveCodecLimits limits)
    {
        _maxTotalBytes = Math.Min(limits.MaxUncompressedBytes, (ulong)MaxAssets * MaxAssetBytes);
        foreach (var asset in assets ?? []) AddRequested(asset);
    }

    internal IReadOnlyCollection<Asset> ImportedAssets => _imported.Values;

    internal Asset Get(string assetId) => _requested.TryGetValue(assetId, out var asset)
        ? asset
        : throw InvalidAsset($"Worksheet image references missing asset {assetId}.");

    internal Asset Import(ImagePart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        var data = memory.ToArray();
        var contentType = NormalizeContentType(part.ContentType);
        ValidateImage(contentType, data, $"Worksheet image part {part.Uri}");
        var digest = Hash(data);
        var id = AssetPrefix + digest;
        if (_requested.TryGetValue(id, out var requested)) return requested;
        if (_imported.TryGetValue(id, out var existing)) return existing;
        if (_imported.Count >= MaxAssets)
            throw new CodecException("spreadsheet_image_asset_budget_exceeded", $"Workbook exceeds the {MaxAssets}-image asset budget.");
        EnsureBudget(data.LongLength);
        var asset = new Asset
        {
            Id = id,
            FileName = $"worksheet-image-{digest[..16]}.{Extension(contentType)}",
            ContentType = contentType,
            Data = ByteString.CopyFrom(data),
            Sha256 = digest,
        };
        _imported.Add(id, asset);
        return asset;
    }

    internal static PartTypeInfo PartTypeFor(string contentType) => NormalizeContentType(contentType) switch
    {
        "image/png" => ImagePartType.Png,
        "image/jpeg" => ImagePartType.Jpeg,
        _ => throw InvalidAsset($"Unsupported worksheet image content type {contentType}."),
    };

    private void AddRequested(Asset source)
    {
        if (_requested.Count >= MaxAssets)
            throw new CodecException("spreadsheet_image_asset_budget_exceeded", $"Workbook exceeds the {MaxAssets}-image asset budget.");
        var contentType = NormalizeContentType(source.ContentType);
        var data = source.Data.ToByteArray();
        ValidateImage(contentType, data, $"Worksheet asset {source.Id}");
        var digest = Hash(data);
        if (!source.Sha256.Equals(digest, StringComparison.OrdinalIgnoreCase) ||
            !source.Id.Equals(AssetPrefix + digest, StringComparison.Ordinal))
            throw InvalidAsset($"Worksheet asset {source.Id} is not content-addressed by its bytes.");
        if (!_requested.TryAdd(source.Id, source.Clone()))
            throw InvalidAsset($"Workbook contains duplicate asset ID {source.Id}.");
        EnsureBudget(data.LongLength);
    }

    private void EnsureBudget(long length)
    {
        if (length is <= 0 or > MaxAssetBytes)
            throw new CodecException("spreadsheet_image_asset_budget_exceeded", $"Worksheet image assets must contain 1 through {MaxAssetBytes} bytes.");
        _totalBytes = checked(_totalBytes + (ulong)length);
        if (_totalBytes > _maxTotalBytes)
            throw new CodecException("spreadsheet_image_asset_budget_exceeded", $"Worksheet image assets exceed the {_maxTotalBytes}-byte budget.");
    }

    private static void ValidateImage(string contentType, byte[] data, string label)
    {
        if (data.Length is 0 or > MaxAssetBytes)
            throw InvalidAsset($"{label} must contain 1 through {MaxAssetBytes} bytes.");
        var valid = contentType switch
        {
            "image/png" => data.AsSpan().StartsWith(Convert.FromHexString("89504E470D0A1A0A")),
            "image/jpeg" => data.Length >= 3 && data[0] == 0xff && data[1] == 0xd8 && data[2] == 0xff,
            _ => false,
        };
        if (!valid) throw InvalidAsset($"{label} bytes do not match a supported PNG or JPEG content type.");
    }

    internal static bool HasSameContentType(ImagePart part, Asset asset) =>
        NormalizeContentType(part.ContentType).Equals(NormalizeContentType(asset.ContentType), StringComparison.Ordinal);

    private static string NormalizeContentType(string value) => value.Equals("image/jpg", StringComparison.OrdinalIgnoreCase)
        ? "image/jpeg"
        : value.ToLowerInvariant();

    private static string Extension(string contentType) => contentType == "image/png" ? "png" : "jpg";
    private static string Hash(byte[] data) => Convert.ToHexString(SHA256.HashData(data)).ToLowerInvariant();
    private static CodecException InvalidAsset(string message) => new("invalid_spreadsheet_image_asset", message);
}
