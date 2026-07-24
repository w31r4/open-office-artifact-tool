using System.IO.Compression;
using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OpenOffice.Artifact.Wire.V1;
using P = DocumentFormat.OpenXml.Presentation;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

internal sealed record PptxOleOfficePackageReplacement(string PartPath, string Sha256, byte[] Data);

// Owns the additive, still deliberately narrow OLE package boundary for
// non-XLSX Office documents. It never authors an OLE shell, relationship,
// preview, or package part. One imported, uniquely-bound DOCX payload may be
// replaced with another validated DOCX having the same content type.
//
// The legacy XLSX-specific PptxOleWorkbookCodec remains a compatibility
// adapter. Keeping this profile separate avoids silently broadening that
// established public wire/API contract while leaving a data-shaped place for
// future audited Office package kinds.
internal static class PptxOleOfficePackageCodec
{
    private const string DocumentKind = "docx";
    private const string DocumentContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    private static readonly HashSet<string> RelationshipNamespaces = new(StringComparer.Ordinal)
    {
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "http://purl.oclc.org/ooxml/officeDocument/relationships",
    };
    private static readonly HashSet<string> PresentationNamespaces = new(StringComparer.Ordinal)
    {
        "http://schemas.openxmlformats.org/presentationml/2006/main",
        "http://purl.oclc.org/ooxml/presentationml/main",
    };

    internal static PptxOleOfficePackageReplacement? PrepareReplacement(
        PresentationOpaqueElement original,
        PresentationOpaqueElement requested,
        PptxAssetCatalog assets,
        EffectiveCodecLimits limits)
    {
        if (original.OleOfficePackage is null || string.IsNullOrWhiteSpace(requested.OleOfficePackage?.ReplacementAssetId)) return null;
        var binding = original.OleOfficePackage;
        var asset = assets.GetOleOfficePackage(requested.OleOfficePackage.ReplacementAssetId);
        if (!asset.ContentType.Equals(binding.ContentType, StringComparison.OrdinalIgnoreCase))
            throw new CodecException("invalid_presentation_asset", "Presentation embedded Office package replacement must retain the source content type.");
        var data = asset.Data.ToByteArray();
        ValidateOfficePackage(binding.Kind, binding.ContentType, data, limits);
        return new PptxOleOfficePackageReplacement(binding.PartPath, asset.Sha256.ToLowerInvariant(), data);
    }

    internal static void Apply(
        SlidePart owner,
        OpenXmlElement source,
        PresentationOleOfficePackage binding,
        PptxOleOfficePackageReplacement replacement)
    {
        var expected = Profile(binding.Kind, binding.ContentType);
        var oleObjects = new[] { source }.Concat(source.Descendants())
            .Where(element => element.LocalName == "oleObj" && PresentationNamespaces.Contains(element.NamespaceUri))
            .ToArray();
        if (oleObjects.Length != 1)
            throw BindingMismatch("Editable OLE Office package owner no longer contains exactly one p:oleObj element.", PartPath(owner));
        var relationshipAttributes = oleObjects[0].GetAttributes()
            .Where(attribute => attribute.LocalName == "id" && RelationshipNamespaces.Contains(attribute.NamespaceUri))
            .ToArray();
        if (relationshipAttributes.Length != 1 || relationshipAttributes[0].Value != binding.RelationshipId)
            throw BindingMismatch("Editable OLE Office package relationship ID no longer matches its source binding.", PartPath(owner));

        OpenXmlPart part;
        try
        {
            part = owner.GetPartById(binding.RelationshipId);
        }
        catch (ArgumentOutOfRangeException exception)
        {
            throw new CodecException("presentation_ole_office_package_binding_mismatch", "Editable OLE Office package relationship no longer resolves to a package part.", PartPath(owner), exception);
        }
        var partPath = PartPath(part);
        if (!partPath.Equals(binding.PartPath, StringComparison.OrdinalIgnoreCase) ||
            !partPath.Equals(replacement.PartPath, StringComparison.OrdinalIgnoreCase) ||
            !part.ContentType.Equals(expected.ContentType, StringComparison.OrdinalIgnoreCase) ||
            !part.RelationshipType.EndsWith("/package", StringComparison.Ordinal))
            throw BindingMismatch("Editable OLE Office package path, content type, or relationship type no longer matches its source binding.", partPath);

        using (var sourceStream = part.GetStream(FileMode.Open, FileAccess.Read))
        using (var memory = new MemoryStream())
        {
            sourceStream.CopyTo(memory);
            if (!Hash(memory.ToArray()).Equals(binding.SourceSha256, StringComparison.OrdinalIgnoreCase))
                throw BindingMismatch("Editable OLE Office package bytes no longer match their source digest.", partPath);
        }
        using var output = part.GetStream(FileMode.Create, FileAccess.Write);
        output.Write(replacement.Data);
    }

    private static void ValidateOfficePackage(string kind, string contentType, byte[] bytes, EffectiveCodecLimits limits)
    {
        var profile = Profile(kind, contentType);
        try
        {
            PackageGuards.ValidateAndCollectOpaque(bytes, limits, profile.OpcProfile, includeSourcePackage: false);
            using var stream = new MemoryStream(bytes, writable: false);
            using var document = WordprocessingDocument.Open(stream, isEditable: false);
            var mainPart = document.MainDocumentPart ??
                throw Invalid("Replacement DOCX has no main document part.");
            var body = mainPart.Document?.Body ??
                throw Invalid("Replacement DOCX has no document body.");
            var paragraphCount = body.Descendants<W.Paragraph>().LongCount();
            if ((ulong)paragraphCount > limits.MaxCells)
                throw new CodecException("presentation_item_budget_exceeded", $"Replacement DOCX has {paragraphCount} paragraphs and exceeds max_cells ({limits.MaxCells}).");
            var validationError = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document).FirstOrDefault();
            if (validationError is not null)
                throw Invalid($"Replacement DOCX fails Office 2021 validation: {validationError.Description}", validationError.Path?.XPath);
        }
        catch (CodecException)
        {
            throw;
        }
        catch (Exception exception) when (exception is OpenXmlPackageException or InvalidDataException or IOException or UnauthorizedAccessException)
        {
            throw new CodecException("invalid_presentation_ole_office_package", $"Replacement {profile.Kind.ToUpperInvariant()} payload is not a readable Office package.", innerException: exception);
        }
    }

    private static OfficePackageProfile Profile(string kind, string contentType)
    {
        if (kind.Equals(DocumentKind, StringComparison.Ordinal) && contentType.Equals(DocumentContentType, StringComparison.OrdinalIgnoreCase))
            return new OfficePackageProfile(DocumentKind, DocumentContentType, OpcPackageProfile.Docx);
        throw new CodecException("invalid_presentation_ole_office_package", "The OLE Office package binding has an unsupported kind or content type.");
    }

    private static CodecException BindingMismatch(string message, string path) =>
        new("presentation_ole_office_package_binding_mismatch", message, path);

    private static CodecException Invalid(string message, string? path = null) =>
        new("invalid_presentation_ole_office_package", message, path);

    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static string Hash(byte[] data) => Convert.ToHexString(SHA256.HashData(data)).ToLowerInvariant();

    private sealed record OfficePackageProfile(string Kind, string ContentType, OpcPackageProfile OpcProfile);
}
