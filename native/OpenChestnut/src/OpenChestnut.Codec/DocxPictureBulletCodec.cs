using System.Globalization;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenOffice.Artifact.Wire.V1;
using V = DocumentFormat.OpenXml.Vml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Owns one deliberately narrow Word picture-numbering profile. The public
// artifact carries semantic image identity and size; numPicBulletId plus OPC
// relationship IDs remain package-local details. Imported DrawingML bullets,
// transformed VML, and other irregular native graphs fail closed.
internal static class DocxPictureBulletCodec
{
    private const long EmuPerPoint = 12_700L;
    private const long MinSizeEmu = 4L * EmuPerPoint;
    private const long MaxSizeEmu = 72L * EmuPerPoint;
    private const string ImageRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    private static readonly XNamespace Wml = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace Vml = "urn:schemas-microsoft-com:vml";
    private static readonly XNamespace Office = "urn:schemas-microsoft-com:office:office";
    private static readonly XNamespace Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    internal static string SemanticKey(DocumentPictureBullet? source)
    {
        if (source is null) return "none";
        var locator = source.SourceCase switch
        {
            DocumentPictureBullet.SourceOneofCase.AssetId => $"asset:{source.AssetId}",
            DocumentPictureBullet.SourceOneofCase.Uri => $"uri:{source.Uri}",
            _ => "missing",
        };
        return string.Join('\0', locator, source.WidthEmu, source.HeightEmu, source.AltText);
    }

    internal static void Validate(
        DocumentPictureBullet? source,
        DocxImageAssetCatalog? assets,
        string label)
    {
        if (source is null) return;
        if (source.WidthEmu is < MinSizeEmu or > MaxSizeEmu ||
            source.HeightEmu is < MinSizeEmu or > MaxSizeEmu)
            throw Invalid($"{label} width and height must be between 4 and 72 points.");
        if (string.IsNullOrWhiteSpace(source.AltText) || source.AltText.Length > 255 || source.AltText.Any(char.IsControl))
            throw Invalid($"{label} alternative text must contain 1 through 255 characters without controls.");

        switch (source.SourceCase)
        {
            case DocumentPictureBullet.SourceOneofCase.AssetId:
                if (source.AssetId.Length > 512 || assets is null)
                    throw Invalid($"{label} requires a bounded image asset catalog reference.");
                var asset = assets.Get(source.AssetId);
                _ = DocxImageAssetCatalog.PartTypeFor(asset.ContentType);
                break;
            case DocumentPictureBullet.SourceOneofCase.Uri:
                if (!TryHttpUri(source.Uri, out _))
                    throw Invalid($"{label} external URI must be absolute http(s), contain at most 4096 characters, and contain no controls.");
                break;
            default:
                throw Invalid($"{label} requires exactly one embedded asset or external URI source.");
        }
    }

    internal static W.NumberingPictureBullet Author(
        NumberingDefinitionsPart part,
        int pictureBulletId,
        DocumentPictureBullet source,
        DocxImageAssetCatalog assets,
        Action<string>? relationshipAdded = null,
        Action<OpenXmlPart>? partAdded = null)
    {
        Validate(source, assets, $"DOCX picture bullet {pictureBulletId}");
        string relationshipId;
        if (source.SourceCase == DocumentPictureBullet.SourceOneofCase.AssetId)
        {
            var asset = assets.Get(source.AssetId);
            var imagePart = part.AddImagePart(DocxImageAssetCatalog.PartTypeFor(asset.ContentType));
            using (var input = new MemoryStream(asset.Data.ToByteArray(), writable: false)) imagePart.FeedData(input);
            relationshipId = part.GetIdOfPart(imagePart);
            partAdded?.Invoke(imagePart);
        }
        else
        {
            _ = TryHttpUri(source.Uri, out var uri);
            relationshipId = part.AddExternalRelationship(ImageRelationshipType, uri!).Id;
        }
        relationshipAdded?.Invoke(relationshipId);

        var shape = new V.Shape(
            new V.ImageData
            {
                RelationshipId = relationshipId,
                Title = source.AltText,
            })
        {
            Id = $"_x0000_i{checked(1025 + pictureBulletId)}",
            Style = $"width:{Points(source.WidthEmu)}pt;height:{Points(source.HeightEmu)}pt",
            Alternate = source.AltText,
            Bullet = true,
        };
        return new W.NumberingPictureBullet(new W.PictureBulletBase(shape))
        {
            NumberingPictureBulletId = pictureBulletId,
        };
    }

    internal static bool TryRead(
        DocxPartContext context,
        XDocument numbering,
        XElement level,
        out DocumentPictureBullet? result)
    {
        result = null;
        var references = level.Elements(Wml + "lvlPicBulletId").Take(2).ToArray();
        if (references.Length == 0) return true;
        if (references.Length != 1 ||
            !int.TryParse(references[0].Attribute(Wml + "val")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var id) ||
            id < 0) return false;
        var definitions = numbering.Root?.Elements(Wml + "numPicBullet")
            .Where(element => int.TryParse(element.Attribute(Wml + "numPicBulletId")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var candidate) && candidate == id)
            .Take(2)
            .ToArray() ?? [];
        if (definitions.Length != 1 || !TryReadDefinition(context, definitions[0], out result))
        {
            result = null;
            return false;
        }
        return true;
    }

    internal static int AuthorSource(
        DocxPartContext context,
        XDocument numbering,
        DocumentPictureBullet source)
    {
        var root = numbering.Root ?? throw Unsupported("The source Numbering part has no root element.");
        var part = context.Owner.NumberingDefinitionsPart ?? throw Unsupported("Picture-bullet edits require a source Numbering part.");
        Validate(source, context.Images, "DOCX picture bullet edit");
        var usedIds = root.Elements(Wml + "numPicBullet")
            .Select(element => int.TryParse(element.Attribute(Wml + "numPicBulletId")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var id) ? id : -1)
            .Where(id => id >= 0)
            .ToHashSet();
        var pictureBulletId = 0;
        while (usedIds.Contains(pictureBulletId)) pictureBulletId = checked(pictureBulletId + 1);
        var authored = Author(
            part,
            pictureBulletId,
            source,
            context.Images ?? throw Unsupported("Picture-bullet edits require an image asset catalog."),
            context.MarkNumberingRelationshipMutated,
            context.MarkPartMutated);
        var element = XElement.Parse(authored.OuterXml, LoadOptions.PreserveWhitespace);
        var firstAbstract = root.Elements(Wml + "abstractNum").FirstOrDefault();
        if (firstAbstract is null) root.Add(element);
        else firstAbstract.AddBeforeSelf(element);
        return pictureBulletId;
    }

    private static bool TryReadDefinition(
        DocxPartContext context,
        XElement definition,
        out DocumentPictureBullet? result)
    {
        result = null;
        var pictures = definition.Elements(Wml + "pict").Take(2).ToArray();
        if (pictures.Length != 1 || definition.Elements(Wml + "drawing").Any() ||
            definition.Elements().Any(element => element.Name != Wml + "pict")) return false;
        var picture = pictures[0];
        var shapes = picture.Elements(Vml + "shape").Take(2).ToArray();
        if (shapes.Length != 1 || picture.Elements().Any(element => element.Name != Vml + "shape")) return false;
        var shape = shapes[0];
        if (!True(shape.Attribute(Office + "bullet")?.Value) ||
            !TryDimensions(shape.Attribute("style")?.Value, out var widthEmu, out var heightEmu)) return false;
        var allowedShapeAttributes = new HashSet<XName>
        {
            "id", "style", "alt", Office + "bullet",
        };
        if (shape.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).Any(attribute => !allowedShapeAttributes.Contains(attribute.Name))) return false;
        var images = shape.Elements(Vml + "imagedata").Take(2).ToArray();
        if (images.Length != 1 || shape.Elements().Any(element => element.Name != Vml + "imagedata")) return false;
        var image = images[0];
        var allowedImageAttributes = new HashSet<XName> { Relationships + "id", Office + "title" };
        if (image.HasElements || image.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).Any(attribute => !allowedImageAttributes.Contains(attribute.Name))) return false;
        var relationshipId = image.Attribute(Relationships + "id")?.Value;
        if (string.IsNullOrWhiteSpace(relationshipId)) return false;
        var shapeAlt = shape.Attribute("alt")?.Value;
        var imageTitle = image.Attribute(Office + "title")?.Value;
        if (!string.IsNullOrEmpty(shapeAlt) && !string.IsNullOrEmpty(imageTitle) && !shapeAlt.Equals(imageTitle, StringComparison.Ordinal)) return false;
        var alt = shapeAlt ?? imageTitle ?? "Picture bullet";
        if (string.IsNullOrWhiteSpace(alt) || alt.Length > 255 || alt.Any(char.IsControl)) return false;

        var part = context.Owner.NumberingDefinitionsPart;
        if (part is null) return false;
        var external = part.ExternalRelationships.Where(item => item.Id == relationshipId).Take(2).ToArray();
        if (external.Length == 1)
        {
            if (!external[0].RelationshipType.Equals(ImageRelationshipType, StringComparison.Ordinal) ||
                !TryHttpUri(external[0].Uri.OriginalString, out _)) return false;
            result = new DocumentPictureBullet
            {
                Uri = external[0].Uri.OriginalString,
                WidthEmu = widthEmu,
                HeightEmu = heightEmu,
                AltText = alt,
            };
            return true;
        }
        if (external.Length != 0) return false;
        try
        {
            if (part.GetPartById(relationshipId) is not ImagePart imagePart) return false;
            var asset = context.Images?.Import(imagePart);
            if (asset is null) return false;
            result = new DocumentPictureBullet
            {
                AssetId = asset.Id,
                WidthEmu = widthEmu,
                HeightEmu = heightEmu,
                AltText = alt,
            };
            return true;
        }
        catch (Exception exception) when (exception is ArgumentOutOfRangeException or KeyNotFoundException or CodecException)
        {
            result = null;
            return false;
        }
    }

    private static bool TryDimensions(string? style, out long widthEmu, out long heightEmu)
    {
        widthEmu = 0;
        heightEmu = 0;
        var values = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        foreach (var declaration in (style ?? string.Empty).Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            var separator = declaration.IndexOf(':');
            if (separator <= 0 || separator == declaration.Length - 1) return false;
            var name = declaration[..separator].Trim();
            var value = declaration[(separator + 1)..].Trim();
            if (name is not ("width" or "height") || !value.EndsWith("pt", StringComparison.OrdinalIgnoreCase) ||
                !decimal.TryParse(value[..^2], NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var points)) return false;
            var emu = decimal.Round(points * EmuPerPoint, 0, MidpointRounding.AwayFromZero);
            if (emu < MinSizeEmu || emu > MaxSizeEmu || !values.TryAdd(name, checked((long)emu))) return false;
        }
        return values.TryGetValue("width", out widthEmu) && values.TryGetValue("height", out heightEmu);
    }

    private static string Points(long emu) =>
        ((decimal)emu / EmuPerPoint).ToString("0.#####", CultureInfo.InvariantCulture);

    private static bool TryHttpUri(string value, out Uri? uri)
    {
        uri = null;
        if (string.IsNullOrWhiteSpace(value) || value.Length > 4096 || value.Any(char.IsControl) ||
            !Uri.TryCreate(value, UriKind.Absolute, out var parsed) || parsed.Scheme is not ("http" or "https")) return false;
        uri = parsed;
        return true;
    }

    private static bool True(string? value) => value is "t" or "true" or "1";
    private static CodecException Invalid(string message) => new("invalid_document_picture_bullet", message, "word/numbering.xml");
    private static CodecException Unsupported(string message) => new("unsupported_document_edit", message, "word/numbering.xml");
}
