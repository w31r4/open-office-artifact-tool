using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml;
using System.Xml.Linq;
using OpenOffice.Artifact.Wire.V1;

namespace OpenOffice.OpenXmlCodec;

internal sealed record EffectiveCodecLimits(
    ulong MaxInputBytes,
    ulong MaxUncompressedBytes,
    uint MaxParts,
    uint MaxSheets,
    ulong MaxCells,
    uint MaxCompressionRatio)
{
    internal static EffectiveCodecLimits From(CodecLimits? limits) => new(
        limits?.MaxInputBytes > 0 ? limits.MaxInputBytes : 64UL * 1024 * 1024,
        limits?.MaxUncompressedBytes > 0 ? limits.MaxUncompressedBytes : 256UL * 1024 * 1024,
        limits?.MaxParts > 0 ? limits.MaxParts : 4_096,
        limits?.MaxSheets > 0 ? limits.MaxSheets : 256,
        limits?.MaxCells > 0 ? limits.MaxCells : 1_000_000,
        limits?.MaxCompressionRatio > 0 ? limits.MaxCompressionRatio : 200);
}

internal static class PackageGuards
{
    private static readonly HashSet<string> OwnedPaths = new(StringComparer.OrdinalIgnoreCase)
    {
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/workbook.xml",
        "xl/_rels/workbook.xml.rels",
    };

    internal static OpaqueOpcGraph ValidateAndCollectOpaque(byte[] bytes, EffectiveCodecLimits limits)
    {
        if ((ulong)bytes.LongLength > limits.MaxInputBytes)
            throw new CodecException("input_budget_exceeded", $"XLSX input has {bytes.LongLength} bytes and exceeds max_input_bytes ({limits.MaxInputBytes}).");

        try
        {
            using var stream = new MemoryStream(bytes, writable: false);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
            if ((uint)archive.Entries.Count > limits.MaxParts)
                throw new CodecException("part_budget_exceeded", $"XLSX package has {archive.Entries.Count} parts and exceeds max_parts ({limits.MaxParts}).");

            ulong totalUncompressed = 0;
            var opaque = new OpaqueOpcGraph();
            var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in archive.Entries)
            {
                ValidateEntryPath(entry.FullName);
                if (!paths.Add(entry.FullName))
                    throw new CodecException("duplicate_part_path", $"XLSX package contains duplicate part path {entry.FullName}.", entry.FullName);
                totalUncompressed = checked(totalUncompressed + (ulong)entry.Length);
                if (totalUncompressed > limits.MaxUncompressedBytes)
                    throw new CodecException("decompression_budget_exceeded", $"XLSX package expands to more than max_uncompressed_bytes ({limits.MaxUncompressedBytes}).", entry.FullName);
                if (entry.Length > 0)
                {
                    var compressed = Math.Max(1L, entry.CompressedLength);
                    var ratio = (ulong)entry.Length / (ulong)compressed;
                    if (ratio > limits.MaxCompressionRatio)
                        throw new CodecException("compression_ratio_exceeded", $"XLSX part {entry.FullName} has compression ratio {ratio}, above max_compression_ratio ({limits.MaxCompressionRatio}).", entry.FullName);
                }

                if (entry.FullName.EndsWith("/", StringComparison.Ordinal)) continue;
                using var partStream = entry.Open();
                using var copy = new MemoryStream();
                partStream.CopyTo(copy);
                var data = copy.ToArray();
                if (entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                    CollectOpaqueRelationships(entry.FullName, data, opaque);
                if (IsOwned(entry.FullName)) continue;
                opaque.Parts.Add(new OpaqueOpcPart
                {
                    Path = entry.FullName,
                    Data = Google.Protobuf.ByteString.CopyFrom(data),
                    Sha256 = Convert.ToHexString(SHA256.HashData(data)).ToLowerInvariant(),
                });
            }
            return opaque;
        }
        catch (CodecException)
        {
            throw;
        }
        catch (InvalidDataException exception)
        {
            throw new CodecException("invalid_opc_package", "XLSX input is not a readable OPC ZIP package.", innerException: exception);
        }
    }

    private static bool IsOwned(string path) => OwnedPaths.Contains(path) || IsWorksheetXml(path);

    private static bool IsWorksheetXml(string path)
    {
        const string prefix = "xl/worksheets/sheet";
        const string suffix = ".xml";
        if (!path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) || !path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)) return false;
        return path[(prefix.Length)..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static void CollectOpaqueRelationships(string relationshipPath, byte[] data, OpaqueOpcGraph opaque)
    {
        try
        {
            using var stream = new MemoryStream(data, writable: false);
            using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
            var document = XDocument.Load(reader, LoadOptions.None);
            var sourcePath = RelationshipSourcePath(relationshipPath);
            foreach (var element in document.Descendants().Where(item => item.Name.LocalName == "Relationship"))
            {
                var relationship = new OpaqueOpcRelationship
                {
                    Id = element.Attribute("Id")?.Value ?? string.Empty,
                    Type = element.Attribute("Type")?.Value ?? string.Empty,
                    Target = element.Attribute("Target")?.Value ?? string.Empty,
                    TargetMode = element.Attribute("TargetMode")?.Value ?? string.Empty,
                    SourcePath = sourcePath,
                };
                if (string.IsNullOrWhiteSpace(relationship.Id) || string.IsNullOrWhiteSpace(relationship.Type) || string.IsNullOrWhiteSpace(relationship.Target))
                    throw new CodecException("invalid_relationship", $"Relationship part {relationshipPath} contains an incomplete relationship.", relationshipPath);
                if (!IsOwnedRelationship(relationship)) opaque.PackageRelationships.Add(relationship);
            }
        }
        catch (CodecException)
        {
            throw;
        }
        catch (XmlException exception)
        {
            throw new CodecException("invalid_relationship_xml", $"Relationship part {relationshipPath} is not valid XML.", relationshipPath, exception);
        }
    }

    private static bool IsOwnedRelationship(OpaqueOpcRelationship relationship)
    {
        if (relationship.TargetMode.Equals("External", StringComparison.OrdinalIgnoreCase)) return false;
        if (relationship.SourcePath.Length == 0)
            return relationship.Type.EndsWith("/officeDocument", StringComparison.Ordinal) && relationship.Target.TrimStart('/').Equals("xl/workbook.xml", StringComparison.OrdinalIgnoreCase);
        return relationship.SourcePath.Equals("xl/workbook.xml", StringComparison.OrdinalIgnoreCase) && relationship.Type.EndsWith("/worksheet", StringComparison.Ordinal);
    }

    private static string RelationshipSourcePath(string relationshipPath)
    {
        if (relationshipPath.Equals("_rels/.rels", StringComparison.OrdinalIgnoreCase)) return string.Empty;
        var marker = relationshipPath.LastIndexOf("/_rels/", StringComparison.OrdinalIgnoreCase);
        if (marker < 0 || !relationshipPath.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            throw new CodecException("invalid_relationship_path", $"Relationship part path {relationshipPath} is invalid.", relationshipPath);
        var directory = relationshipPath[..marker];
        var fileName = relationshipPath[(marker + "/_rels/".Length)..^".rels".Length];
        return directory.Length == 0 ? fileName : $"{directory}/{fileName}";
    }

    private static void ValidateEntryPath(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || path.StartsWith("/", StringComparison.Ordinal) || path.Contains('\\'))
            throw new CodecException("unsafe_part_path", $"XLSX package contains unsafe part path {path}.", path);
        var segments = path.Split('/', StringSplitOptions.RemoveEmptyEntries);
        if (segments.Any(segment => segment is "." or ".."))
            throw new CodecException("unsafe_part_path", $"XLSX package contains unsafe part path {path}.", path);
    }
}
