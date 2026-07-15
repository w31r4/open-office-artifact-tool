using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns durable picture-bullet asset identity. Open Packaging Convention part
// names and relationship IDs intentionally remain outside the wire contract.
internal sealed class PptxAssetCatalog
{
    private const int MaxAssets = 1_024;
    private const int MaxAssetBytes = 16 * 1024 * 1024;
    private const string PictureAssetPrefix = "asset/presentation/picture-bullet/";
    private const string OleWorkbookAssetPrefix = "asset/presentation/ole-workbook/";
    private const string SpreadsheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    private readonly Dictionary<string, Asset> _assets = new(StringComparer.Ordinal);
    private readonly Dictionary<string, ImagePart> _partByAssetId = new(StringComparer.Ordinal);
    private readonly Dictionary<string, Asset> _imported = new(StringComparer.Ordinal);
    private readonly ulong _maxTotalBytes;
    private ulong _totalBytes;

    internal PptxAssetCatalog(IEnumerable<Asset>? assets, EffectiveCodecLimits limits)
    {
        _maxTotalBytes = Math.Min(limits.MaxUncompressedBytes, (ulong)MaxAssets * MaxAssetBytes);
        foreach (var asset in assets ?? []) AddRequested(asset);
    }

    internal IReadOnlyCollection<Asset> ImportedAssets => _imported.Values;

    internal Asset Get(string assetId) => _assets.TryGetValue(assetId, out var asset)
        ? asset.Id.StartsWith(PictureAssetPrefix, StringComparison.Ordinal)
            ? asset
            : throw new CodecException("invalid_presentation_asset", $"Presentation picture bullet references non-image asset {assetId}.")
        : throw new CodecException("invalid_presentation_asset", $"Presentation picture bullet references missing asset {assetId}.");

    internal Asset GetOleWorkbook(string assetId) => _assets.TryGetValue(assetId, out var asset) &&
        asset.Id.StartsWith(OleWorkbookAssetPrefix, StringComparison.Ordinal)
            ? asset
            : throw new CodecException("invalid_presentation_asset", $"Presentation OLE workbook references missing asset {assetId}.");

    internal Asset Import(ImagePart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        var data = memory.ToArray();
        var contentType = NormalizeContentType(part.ContentType);
        ValidateImage(contentType, data, $"Presentation image part {part.Uri}");
        var digest = Hash(data);
        var id = PictureAssetPrefix + digest;
        if (_assets.TryGetValue(id, out var requested))
        {
            _partByAssetId.TryAdd(id, part);
            return requested;
        }
        if (!_imported.TryGetValue(id, out var asset))
        {
            if (_imported.Count >= MaxAssets)
                throw new CodecException("presentation_asset_budget_exceeded", $"Presentation exceeds the {MaxAssets}-asset budget.");
            EnsureBudget(data.LongLength);
            asset = new Asset
            {
                Id = id,
                FileName = $"picture-bullet-{digest[..16]}.{Extension(contentType)}",
                ContentType = contentType,
                Data = ByteString.CopyFrom(data),
                Sha256 = digest,
            };
            _imported.Add(id, asset);
        }
        _partByAssetId.TryAdd(id, part);
        return asset;
    }

    internal ImagePart? ExistingPart(string assetId) => _partByAssetId.GetValueOrDefault(assetId);

    internal void RegisterPart(string assetId, ImagePart part) => _partByAssetId.TryAdd(assetId, part);

    internal void IndexExistingParts(IEnumerable<ImagePart> parts)
    {
        foreach (var part in parts.Distinct())
        {
            try
            {
                using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
                using var memory = new MemoryStream();
                stream.CopyTo(memory);
                var data = memory.ToArray();
                var id = PictureAssetPrefix + Hash(data);
                if (_assets.TryGetValue(id, out var asset) && part.ContentType.Equals(asset.ContentType, StringComparison.OrdinalIgnoreCase))
                    _partByAssetId.TryAdd(id, part);
            }
            catch (IOException)
            {
                // Opaque/unreadable image parts are guarded elsewhere and are
                // never selected as modeled assets.
            }
        }
    }

    internal static PartTypeInfo ImagePartTypeFor(string contentType) => NormalizeContentType(contentType) switch
    {
        "image/png" => ImagePartType.Png,
        "image/jpeg" => ImagePartType.Jpeg,
        "image/gif" => ImagePartType.Gif,
        "image/svg+xml" => ImagePartType.Svg,
        _ => throw new CodecException("invalid_presentation_asset", $"Unsupported presentation image content type {contentType}."),
    };

    private void AddRequested(Asset source)
    {
        if (_assets.Count >= MaxAssets)
            throw new CodecException("presentation_asset_budget_exceeded", $"Presentation exceeds the {MaxAssets}-asset budget.");
        var contentType = NormalizeContentType(source.ContentType);
        var data = source.Data.ToByteArray();
        var isPicture = source.Id.StartsWith(PictureAssetPrefix, StringComparison.Ordinal);
        var isOleWorkbook = source.Id.StartsWith(OleWorkbookAssetPrefix, StringComparison.Ordinal);
        if (isPicture) ValidateImage(contentType, data, $"Presentation asset {source.Id}");
        else if (isOleWorkbook) ValidateOleWorkbook(contentType, data, $"Presentation asset {source.Id}");
        else throw new CodecException("invalid_presentation_asset", $"Presentation asset ID {source.Id} has an unsupported purpose prefix.");
        var digest = Hash(data);
        if (!source.Sha256.Equals(digest, StringComparison.OrdinalIgnoreCase))
            throw new CodecException("invalid_presentation_asset", $"Presentation asset {source.Id} does not match its SHA-256 digest.");
        var expectedId = (isPicture ? PictureAssetPrefix : OleWorkbookAssetPrefix) + digest;
        if (!source.Id.Equals(expectedId, StringComparison.Ordinal))
            throw new CodecException("invalid_presentation_asset", $"Presentation asset {source.Id} is not content-addressed by its bytes.");
        if (!_assets.TryAdd(source.Id, source.Clone()))
            throw new CodecException("invalid_presentation_asset", $"Presentation contains duplicate asset ID {source.Id}.");
        EnsureBudget(data.LongLength);
    }

    private void EnsureBudget(long length)
    {
        if (length <= 0 || length > MaxAssetBytes)
            throw new CodecException("presentation_asset_budget_exceeded", $"Presentation picture-bullet assets must contain 1 through {MaxAssetBytes} bytes.");
        _totalBytes = checked(_totalBytes + (ulong)length);
        if (_totalBytes > _maxTotalBytes)
            throw new CodecException("presentation_asset_budget_exceeded", $"Presentation picture-bullet assets exceed the {_maxTotalBytes}-byte budget.");
    }

    private static void ValidateImage(string contentType, byte[] data, string label)
    {
        if (data.Length is 0 or > MaxAssetBytes)
            throw new CodecException("invalid_presentation_asset", $"{label} must contain 1 through {MaxAssetBytes} bytes.");
        var valid = contentType switch
        {
            "image/png" => data.AsSpan().StartsWith(Convert.FromHexString("89504E470D0A1A0A")),
            "image/jpeg" => data.Length >= 3 && data[0] == 0xff && data[1] == 0xd8 && data[2] == 0xff,
            "image/gif" => data.Length >= 6 && Encoding.ASCII.GetString(data, 0, 6) is "GIF87a" or "GIF89a",
            "image/svg+xml" => IsSafeSvg(data),
            _ => false,
        };
        if (!valid) throw new CodecException("invalid_presentation_asset", $"{label} bytes do not match a supported PNG, JPEG, GIF, or safe SVG content type.");
    }

    private static void ValidateOleWorkbook(string contentType, byte[] data, string label)
    {
        if (!contentType.Equals(SpreadsheetContentType, StringComparison.Ordinal))
            throw new CodecException("invalid_presentation_asset", $"{label} must use the XLSX workbook content type.");
        if (data.Length is 0 or > MaxAssetBytes || data.Length < 4 ||
            data[0] != 0x50 || data[1] != 0x4b || data[2] != 0x03 || data[3] != 0x04)
            throw new CodecException("invalid_presentation_asset", $"{label} must contain 1 through {MaxAssetBytes} bytes of an OPC ZIP package.");
    }

    private static bool IsSafeSvg(byte[] data)
    {
        try
        {
            using var stream = new MemoryStream(data, writable: false);
            using var reader = XmlReader.Create(stream, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = MaxAssetBytes,
                IgnoreComments = true,
            });
            var document = XDocument.Load(reader, LoadOptions.None);
            if (document.Root?.Name.LocalName != "svg") return false;
            if (document.DescendantNodes().OfType<XProcessingInstruction>().Any()) return false;
            foreach (var element in document.Root.DescendantsAndSelf())
            {
                if (element.Name.LocalName is "script" or "foreignObject") return false;
                if (element.Name.LocalName == "style" && UnsafeCss(element.Value)) return false;
                foreach (var attribute in element.Attributes())
                {
                    if (attribute.Name.LocalName.StartsWith("on", StringComparison.OrdinalIgnoreCase)) return false;
                    if (UnsafeCss(attribute.Value)) return false;
                    if (attribute.Name.LocalName != "href") continue;
                    var target = attribute.Value.Trim();
                    if (target.Length > 0 && !target.StartsWith('#') &&
                        !target.StartsWith("data:image/png;base64,", StringComparison.OrdinalIgnoreCase) &&
                        !target.StartsWith("data:image/jpeg;base64,", StringComparison.OrdinalIgnoreCase) &&
                        !target.StartsWith("data:image/gif;base64,", StringComparison.OrdinalIgnoreCase)) return false;
                }
            }
            return true;
        }
        catch (XmlException)
        {
            return false;
        }
    }

    private static bool UnsafeCss(string value)
    {
        if (value.Contains("@import", StringComparison.OrdinalIgnoreCase)) return true;
        var offset = 0;
        while (value.IndexOf("url(", offset, StringComparison.OrdinalIgnoreCase) is var start && start >= 0)
        {
            var end = value.IndexOf(')', start + 4);
            if (end < 0) return true;
            var target = value[(start + 4)..end].Trim().Trim('\'', '"');
            if (target.Length > 0 && !target.StartsWith('#') &&
                !target.StartsWith("data:image/png;base64,", StringComparison.OrdinalIgnoreCase) &&
                !target.StartsWith("data:image/jpeg;base64,", StringComparison.OrdinalIgnoreCase) &&
                !target.StartsWith("data:image/gif;base64,", StringComparison.OrdinalIgnoreCase)) return true;
            offset = end + 1;
        }
        return false;
    }

    private static string NormalizeContentType(string value) => value.Equals("image/jpg", StringComparison.OrdinalIgnoreCase)
        ? "image/jpeg"
        : value.ToLowerInvariant();

    private static string Extension(string contentType) => contentType switch
    {
        "image/png" => "png",
        "image/jpeg" => "jpg",
        "image/gif" => "gif",
        "image/svg+xml" => "svg",
        _ => "bin",
    };

    private static string Hash(byte[] data) => Convert.ToHexString(SHA256.HashData(data)).ToLowerInvariant();
}
