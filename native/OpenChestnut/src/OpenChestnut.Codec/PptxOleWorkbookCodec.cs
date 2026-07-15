using System.IO.Compression;
using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OpenOffice.Artifact.Wire.V1;
using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OpenChestnut.Codec;

internal sealed record PptxOleWorkbookReplacement(string PartPath, string Sha256, byte[] Data);

// Owns the deliberately narrow editable boundary for an existing OLE-backed
// XLSX package. It never creates OLE markup, relationships, preview images, or
// package parts; it replaces bytes in one hash-bound EmbeddedPackagePart.
internal static class PptxOleWorkbookCodec
{
    private const string SpreadsheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
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

    internal static PptxOleWorkbookReplacement? PrepareReplacement(
        PresentationOpaqueElement original,
        PresentationOpaqueElement requested,
        PptxAssetCatalog assets,
        EffectiveCodecLimits limits)
    {
        if (original.OleWorkbook is null || string.IsNullOrWhiteSpace(requested.OleWorkbook?.ReplacementAssetId)) return null;
        var asset = assets.GetOleWorkbook(requested.OleWorkbook.ReplacementAssetId);
        var data = asset.Data.ToByteArray();
        ValidateWorkbook(data, limits);
        return new PptxOleWorkbookReplacement(original.OleWorkbook.PartPath, asset.Sha256.ToLowerInvariant(), data);
    }

    internal static void Apply(
        SlidePart owner,
        OpenXmlElement source,
        PresentationOleWorkbook binding,
        PptxOleWorkbookReplacement replacement)
    {
        var oleObjects = new[] { source }.Concat(source.Descendants())
            .Where(element => element.LocalName == "oleObj" && PresentationNamespaces.Contains(element.NamespaceUri))
            .ToArray();
        if (oleObjects.Length != 1)
            throw new CodecException("presentation_ole_workbook_binding_mismatch", "Editable OLE workbook owner no longer contains exactly one p:oleObj element.", PartPath(owner));
        var relationshipAttributes = oleObjects[0].GetAttributes()
            .Where(attribute => attribute.LocalName == "id" && RelationshipNamespaces.Contains(attribute.NamespaceUri))
            .ToArray();
        if (relationshipAttributes.Length != 1 || relationshipAttributes[0].Value != binding.RelationshipId)
            throw new CodecException("presentation_ole_workbook_binding_mismatch", "Editable OLE workbook relationship ID no longer matches its source binding.", PartPath(owner));

        OpenXmlPart part;
        try
        {
            part = owner.GetPartById(binding.RelationshipId);
        }
        catch (ArgumentOutOfRangeException exception)
        {
            throw new CodecException("presentation_ole_workbook_binding_mismatch", "Editable OLE workbook relationship no longer resolves to a package part.", PartPath(owner), exception);
        }
        var partPath = PartPath(part);
        if (!partPath.Equals(binding.PartPath, StringComparison.OrdinalIgnoreCase) ||
            !partPath.Equals(replacement.PartPath, StringComparison.OrdinalIgnoreCase) ||
            !part.ContentType.Equals(SpreadsheetContentType, StringComparison.OrdinalIgnoreCase) ||
            !part.RelationshipType.EndsWith("/package", StringComparison.Ordinal))
            throw new CodecException("presentation_ole_workbook_binding_mismatch", "Editable OLE workbook part path, content type, or relationship type no longer matches its source binding.", partPath);

        using (var sourceStream = part.GetStream(FileMode.Open, FileAccess.Read))
        using (var memory = new MemoryStream())
        {
            sourceStream.CopyTo(memory);
            if (!Hash(memory.ToArray()).Equals(binding.SourceSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_ole_workbook_binding_mismatch", "Editable OLE workbook bytes no longer match their source digest.", partPath);
        }
        using var output = part.GetStream(FileMode.Create, FileAccess.Write);
        output.Write(replacement.Data);
    }

    private static void ValidateWorkbook(byte[] bytes, EffectiveCodecLimits limits)
    {
        try
        {
            PackageGuards.ValidateAndCollectOpaque(bytes, limits, OpcPackageProfile.Xlsx, includeSourcePackage: false);
            using var stream = new MemoryStream(bytes, writable: false);
            using var workbook = SpreadsheetDocument.Open(stream, isEditable: false);
            var workbookPart = workbook.WorkbookPart ??
                throw new CodecException("invalid_presentation_ole_workbook", "Replacement XLSX has no workbook part.");
            var root = workbookPart.Workbook ??
                throw new CodecException("invalid_presentation_ole_workbook", "Replacement XLSX has no workbook root.");
            var sheets = root.Sheets?.Elements<S.Sheet>().ToArray() ?? [];
            if (sheets.Length == 0)
                throw new CodecException("invalid_presentation_ole_workbook", "Replacement XLSX must contain at least one worksheet.");
            if ((uint)sheets.Length > limits.MaxSheets)
                throw new CodecException("slide_budget_exceeded", $"Replacement XLSX has {sheets.Length} sheets and exceeds max_sheets ({limits.MaxSheets}).");
            ulong cells = 0;
            foreach (var worksheetPart in workbookPart.WorksheetParts)
            {
                cells = checked(cells + (ulong)(worksheetPart.Worksheet?.Descendants<S.Cell>().Count() ?? 0));
                if (cells > limits.MaxCells)
                    throw new CodecException("presentation_item_budget_exceeded", $"Replacement XLSX exceeds max_cells ({limits.MaxCells}).");
            }
            var validationError = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(workbook).FirstOrDefault();
            if (validationError is not null)
                throw new CodecException("invalid_presentation_ole_workbook", $"Replacement XLSX fails Office 2021 validation: {validationError.Description}", validationError.Path?.XPath);
        }
        catch (CodecException)
        {
            throw;
        }
        catch (Exception exception) when (exception is OpenXmlPackageException or InvalidDataException or IOException or UnauthorizedAccessException)
        {
            throw new CodecException("invalid_presentation_ole_workbook", "Replacement OLE payload is not a readable XLSX package.", innerException: exception);
        }
    }

    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static string Hash(byte[] data) => Convert.ToHexString(SHA256.HashData(data)).ToLowerInvariant();
}
