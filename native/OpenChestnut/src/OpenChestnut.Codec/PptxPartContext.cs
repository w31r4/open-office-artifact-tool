using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OpenChestnut.Codec;

// Owns relationships whose source is one PresentationML part. Asset identity
// belongs to the shared catalog; relationship IDs remain local to the slide,
// master, or layout owner part.
internal sealed class PptxPartContext
{
    private const string ImageRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    private readonly HashSet<string> _addedRelationshipIds = new(StringComparer.Ordinal);
    private readonly HashSet<string> _addedPartPaths = new(StringComparer.OrdinalIgnoreCase);
    private readonly Func<PartTypeInfo, ImagePart> _addImagePart;

    internal PptxPartContext(
        OpenXmlPart owner,
        IReadOnlyDictionary<string, string> slideIdByPartPath,
        IReadOnlyDictionary<string, SlidePart>? slidePartById = null,
        PptxAssetCatalog? assets = null,
        PptxCustomShowCatalog? customShows = null) : this(
            owner,
            owner switch
            {
                SlidePart slide => type => slide.AddImagePart(type),
                SlideMasterPart master => type => master.AddImagePart(type),
                SlideLayoutPart layout => type => layout.AddImagePart(type),
                _ => throw new ArgumentException($"Unsupported PresentationML relationship owner {owner.GetType().Name}.", nameof(owner)),
            },
            slideIdByPartPath,
            slidePartById,
            assets,
            customShows)
    {
    }

    private PptxPartContext(
        OpenXmlPart owner,
        Func<PartTypeInfo, ImagePart> addImagePart,
        IReadOnlyDictionary<string, string> slideIdByPartPath,
        IReadOnlyDictionary<string, SlidePart>? slidePartById,
        PptxAssetCatalog? assets,
        PptxCustomShowCatalog? customShows)
    {
        Owner = owner;
        _addImagePart = addImagePart;
        SlideIdByPartPath = slideIdByPartPath;
        SlidePartById = slidePartById ?? new Dictionary<string, SlidePart>(StringComparer.Ordinal);
        Assets = assets;
        CustomShows = customShows ?? PptxCustomShowCatalog.Empty;
    }

    internal OpenXmlPart Owner { get; }
    internal IReadOnlyDictionary<string, string> SlideIdByPartPath { get; }
    internal IReadOnlyDictionary<string, SlidePart> SlidePartById { get; }
    internal PptxAssetCatalog? Assets { get; }
    internal PptxCustomShowCatalog CustomShows { get; }
    internal bool RelationshipsChanged => _addedRelationshipIds.Count > 0;
    internal IReadOnlyCollection<string> AddedRelationshipIds => _addedRelationshipIds;
    internal IReadOnlyCollection<string> AddedPartPaths => _addedPartPaths;

    internal string AddExternalHyperlink(string uri)
    {
        var existing = Owner.HyperlinkRelationships.FirstOrDefault(relationship =>
            relationship.IsExternal && relationship.Uri.OriginalString.Equals(uri, StringComparison.Ordinal));
        if (existing is not null) return existing.Id;
        return Track(Owner.AddHyperlinkRelationship(new Uri(uri, UriKind.Absolute), true).Id);
    }

    internal string AddSlide(string slideId)
    {
        if (!SlidePartById.TryGetValue(slideId, out var target))
            throw new CodecException("invalid_presentation_hyperlink", $"Presentation run hyperlink references missing slide {slideId}.");
        var existing = Owner.Parts.FirstOrDefault(pair => ReferenceEquals(pair.OpenXmlPart, target));
        if (existing.OpenXmlPart is not null) return existing.RelationshipId;
        Owner.AddPart(target);
        return Track(Owner.GetIdOfPart(target));
    }

    internal bool TryReadPicture(A.PictureBullet source, out PresentationPictureBullet picture)
    {
        picture = new PresentationPictureBullet();
        if (Assets is null || source.ChildElements.Count != 1 || source.GetFirstChild<A.Blip>() is not { } blip ||
            blip.ChildElements.Count > 0 || blip.CompressionState is not null) return false;
        var embed = blip.Embed?.Value ?? string.Empty;
        var link = blip.Link?.Value ?? string.Empty;
        if ((embed.Length == 0) == (link.Length == 0)) return false;
        if (embed.Length > 0)
        {
            try
            {
                if (Owner.GetPartById(embed) is not ImagePart imagePart) return false;
                picture.AssetId = Assets.Import(imagePart).Id;
                return true;
            }
            catch (Exception error) when (error is ArgumentOutOfRangeException or CodecException)
            {
                return false;
            }
        }
        var relationship = Owner.ExternalRelationships.FirstOrDefault(item => item.Id == link && item.RelationshipType.EndsWith("/image", StringComparison.Ordinal));
        if (relationship is null) return false;
        try
        {
            picture.Uri = ValidatePictureUri(relationship.Uri.OriginalString);
            return true;
        }
        catch (CodecException)
        {
            picture = new PresentationPictureBullet();
            return false;
        }
    }

    internal A.PictureBullet BuildPicture(PresentationPictureBullet picture)
    {
        var blip = new A.Blip();
        switch (picture.SourceCase)
        {
            case PresentationPictureBullet.SourceOneofCase.AssetId:
                blip.Embed = AddEmbeddedPicture(picture.AssetId);
                break;
            case PresentationPictureBullet.SourceOneofCase.Uri:
                blip.Link = AddExternalPicture(picture.Uri);
                break;
            default:
                throw InvalidPicture("Presentation picture bullet requires exactly one source.");
        }
        return new A.PictureBullet(blip);
    }

    internal static void ValidatePicture(PresentationPictureBullet? picture)
    {
        if (picture is null) throw InvalidPicture("Presentation picture bullet payload is missing.");
        switch (picture.SourceCase)
        {
            case PresentationPictureBullet.SourceOneofCase.AssetId:
                if (string.IsNullOrWhiteSpace(picture.AssetId) || picture.AssetId.Length > 512)
                    throw InvalidPicture("Presentation picture bullet asset ID must contain 1 through 512 characters.");
                break;
            case PresentationPictureBullet.SourceOneofCase.Uri:
                ValidatePictureUri(picture.Uri);
                break;
            default:
                throw InvalidPicture("Presentation picture bullet requires exactly one source.");
        }
    }

    internal Asset ReadEmbeddedPicture(string relationshipId)
    {
        if (Assets is null) throw InvalidPicture("Presentation image import requires an asset catalog.");
        try
        {
            if (Owner.GetPartById(relationshipId) is not ImagePart imagePart)
                throw InvalidPicture($"Presentation image relationship {relationshipId} does not resolve to an image part.");
            return Assets.Import(imagePart);
        }
        catch (ArgumentOutOfRangeException)
        {
            throw InvalidPicture($"Presentation image relationship {relationshipId} is missing.");
        }
    }

    internal string AddEmbeddedPicture(string assetId)
    {
        if (Assets is null) throw InvalidPicture("Presentation picture authoring requires an asset catalog.");
        var asset = Assets.Get(assetId);
        var existingOwnerPart = Owner.Parts.Select(pair => pair.OpenXmlPart).OfType<ImagePart>().FirstOrDefault(part => PartMatches(part, asset));
        if (existingOwnerPart is not null)
        {
            Assets.RegisterPart(assetId, existingOwnerPart);
            return Owner.GetIdOfPart(existingOwnerPart);
        }
        if (Assets.ExistingPart(assetId) is { } shared)
        {
            Owner.AddPart(shared);
            return Track(Owner.GetIdOfPart(shared));
        }
        var part = _addImagePart(PptxAssetCatalog.ImagePartTypeFor(asset.ContentType));
        using (var source = new MemoryStream(asset.Data.ToByteArray(), writable: false)) part.FeedData(source);
        Assets.RegisterPart(assetId, part);
        _addedPartPaths.Add(part.Uri.OriginalString.TrimStart('/'));
        return Track(Owner.GetIdOfPart(part));
    }

    private string AddExternalPicture(string value)
    {
        var uri = ValidatePictureUri(value);
        var existing = Owner.ExternalRelationships.FirstOrDefault(relationship =>
            relationship.RelationshipType.EndsWith("/image", StringComparison.Ordinal) &&
            relationship.Uri.OriginalString.Equals(uri, StringComparison.Ordinal));
        if (existing is not null) return existing.Id;
        return Track(Owner.AddExternalRelationship(ImageRelationshipType, new Uri(uri, UriKind.Absolute)).Id);
    }

    private static bool PartMatches(ImagePart part, Asset asset)
    {
        if (!part.ContentType.Equals(asset.ContentType, StringComparison.OrdinalIgnoreCase)) return false;
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var sha = System.Security.Cryptography.SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(stream)).Equals(asset.Sha256, StringComparison.OrdinalIgnoreCase);
    }

    private string Track(string relationshipId)
    {
        _addedRelationshipIds.Add(relationshipId);
        return relationshipId;
    }

    private static string ValidatePictureUri(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 4_096 || value.Any(char.IsControl) ||
            !Uri.TryCreate(value, UriKind.Absolute, out var uri) ||
            uri.Scheme is not ("http" or "https"))
            throw InvalidPicture("Presentation picture bullet URI must be an absolute http(s) URI of at most 4096 characters without controls.");
        return value;
    }

    private static CodecException InvalidPicture(string message) => new("invalid_presentation_asset", message);
}
