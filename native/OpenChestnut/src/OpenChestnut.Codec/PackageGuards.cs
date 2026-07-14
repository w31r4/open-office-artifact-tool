using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml;
using System.Xml.Linq;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

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

internal sealed class OpcPackageProfile
{
    private readonly Func<string, bool> _ownsPath;
    private readonly Func<OpaqueOpcRelationship, bool> _ownsRelationship;

    private OpcPackageProfile(
        string format,
        Func<string, bool> ownsPath,
        Func<OpaqueOpcRelationship, bool> ownsRelationship)
    {
        Format = format;
        _ownsPath = ownsPath;
        _ownsRelationship = ownsRelationship;
    }

    internal string Format { get; }
    internal bool OwnsPath(string path) => _ownsPath(path);
    internal bool OwnsRelationship(OpaqueOpcRelationship relationship) => _ownsRelationship(relationship);

    private static readonly HashSet<string> CommonOwnedPaths = new(StringComparer.OrdinalIgnoreCase)
    {
        "[Content_Types].xml",
        "_rels/.rels",
    };

    internal static OpcPackageProfile Xlsx { get; } = new(
        "XLSX",
        path => CommonOwnedPaths.Contains(path) ||
                path.Equals("xl/workbook.xml", StringComparison.OrdinalIgnoreCase) ||
                path.Equals("xl/_rels/workbook.xml.rels", StringComparison.OrdinalIgnoreCase) ||
                path.Equals("xl/styles.xml", StringComparison.OrdinalIgnoreCase) ||
                IsNumberedXml(path, "xl/worksheets/sheet"),
        relationship =>
        {
            if (relationship.TargetMode.Equals("External", StringComparison.OrdinalIgnoreCase)) return false;
            if (relationship.SourcePath.Length == 0)
                return relationship.Type.EndsWith("/officeDocument", StringComparison.Ordinal) &&
                       relationship.Target.TrimStart('/').Equals("xl/workbook.xml", StringComparison.OrdinalIgnoreCase);
            return relationship.SourcePath.Equals("xl/workbook.xml", StringComparison.OrdinalIgnoreCase) &&
                   (relationship.Type.EndsWith("/worksheet", StringComparison.Ordinal) ||
                    relationship.Type.EndsWith("/styles", StringComparison.Ordinal));
        });

    internal static OpcPackageProfile Docx { get; } = new(
        "DOCX",
        path => CommonOwnedPaths.Contains(path) ||
                path.Equals("word/document.xml", StringComparison.OrdinalIgnoreCase) ||
                path.Equals("word/_rels/document.xml.rels", StringComparison.OrdinalIgnoreCase),
        relationship =>
        {
            if (relationship.TargetMode.Equals("External", StringComparison.OrdinalIgnoreCase)) return false;
            return relationship.SourcePath.Length == 0 &&
                   relationship.Type.EndsWith("/officeDocument", StringComparison.Ordinal) &&
                   relationship.Target.TrimStart('/').Equals("word/document.xml", StringComparison.OrdinalIgnoreCase);
        });

    internal static OpcPackageProfile Pptx { get; } = new(
        "PPTX",
        path => CommonOwnedPaths.Contains(path) ||
                path.Equals("ppt/presentation.xml", StringComparison.OrdinalIgnoreCase) ||
                path.Equals("ppt/_rels/presentation.xml.rels", StringComparison.OrdinalIgnoreCase) ||
                IsNumberedXml(path, "ppt/slides/slide") ||
                IsNumberedRelationshipXml(path, "ppt/slides/_rels/slide") ||
                IsNumberedXml(path, "ppt/slideMasters/slideMaster") ||
                IsNumberedRelationshipXml(path, "ppt/slideMasters/_rels/slideMaster") ||
                IsNumberedXml(path, "ppt/slideLayouts/slideLayout") ||
                IsNumberedRelationshipXml(path, "ppt/slideLayouts/_rels/slideLayout"),
        relationship =>
        {
            if (relationship.TargetMode.Equals("External", StringComparison.OrdinalIgnoreCase)) return false;
            if (relationship.SourcePath.Length == 0)
                return relationship.Type.EndsWith("/officeDocument", StringComparison.Ordinal) &&
                       relationship.Target.TrimStart('/').Equals("ppt/presentation.xml", StringComparison.OrdinalIgnoreCase);
            return relationship.SourcePath.Equals("ppt/presentation.xml", StringComparison.OrdinalIgnoreCase) &&
                   relationship.Type.EndsWith("/slide", StringComparison.Ordinal);
        });

    private static bool IsNumberedXml(string path, string prefix)
    {
        const string suffix = ".xml";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsNumberedRelationshipXml(string path, string prefix)
    {
        const string suffix = ".xml.rels";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }
}

internal static class PackageGuards
{
    internal static OpaqueOpcGraph ValidateAndCollectOpaque(byte[] bytes, EffectiveCodecLimits limits, OpcPackageProfile profile, bool includeSourcePackage = true)
    {
        if ((ulong)bytes.LongLength > limits.MaxInputBytes)
            throw new CodecException("input_budget_exceeded", $"{profile.Format} input has {bytes.LongLength} bytes and exceeds max_input_bytes ({limits.MaxInputBytes}).");

        try
        {
            using var stream = new MemoryStream(bytes, writable: false);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
            if ((uint)archive.Entries.Count > limits.MaxParts)
                throw new CodecException("part_budget_exceeded", $"{profile.Format} package has {archive.Entries.Count} parts and exceeds max_parts ({limits.MaxParts}).");

            ulong totalUncompressed = 0;
            var opaque = new OpaqueOpcGraph();
            var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in archive.Entries)
            {
                ValidateEntryPath(entry.FullName, profile);
                if (!paths.Add(entry.FullName))
                    throw new CodecException("duplicate_part_path", $"{profile.Format} package contains duplicate part path {entry.FullName}.", entry.FullName);
                totalUncompressed = checked(totalUncompressed + (ulong)entry.Length);
                if (totalUncompressed > limits.MaxUncompressedBytes)
                    throw new CodecException("decompression_budget_exceeded", $"{profile.Format} package expands to more than max_uncompressed_bytes ({limits.MaxUncompressedBytes}).", entry.FullName);
                if (entry.Length > 0)
                {
                    var compressed = Math.Max(1L, entry.CompressedLength);
                    var ratio = (ulong)entry.Length / (ulong)compressed;
                    if (ratio > limits.MaxCompressionRatio)
                        throw new CodecException("compression_ratio_exceeded", $"{profile.Format} part {entry.FullName} has compression ratio {ratio}, above max_compression_ratio ({limits.MaxCompressionRatio}).", entry.FullName);
                }

                if (entry.FullName.EndsWith("/", StringComparison.Ordinal)) continue;
                using var partStream = entry.Open();
                using var copy = new MemoryStream();
                partStream.CopyTo(copy);
                var data = copy.ToArray();
                if (entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                    CollectOpaqueRelationships(entry.FullName, data, opaque, profile);
                if (profile.OwnsPath(entry.FullName)) continue;
                opaque.Parts.Add(new OpaqueOpcPart
                {
                    Path = entry.FullName,
                    Sha256 = Convert.ToHexString(SHA256.HashData(data)).ToLowerInvariant(),
                });
            }
            if (includeSourcePackage)
            {
                var packageHash = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
                opaque.SourcePackage = new SourcePackageSnapshot
                {
                    Data = Google.Protobuf.ByteString.CopyFrom(bytes),
                    Sha256 = packageHash,
                };
            }
            return opaque;
        }
        catch (CodecException)
        {
            throw;
        }
        catch (InvalidDataException exception)
        {
            throw new CodecException("invalid_opc_package", $"{profile.Format} input is not a readable OPC ZIP package.", innerException: exception);
        }
    }

    internal static byte[] ValidateSourcePackage(OpaqueOpcGraph opaque, SourceIdentity? source, EffectiveCodecLimits limits, OpcPackageProfile profile)
    {
        if (opaque.SourcePackage is null || opaque.SourcePackage.Data.IsEmpty)
            throw new CodecException("missing_source_package", "Opaque OPC content cannot be preserved because its source package snapshot is missing.");
        var bytes = opaque.SourcePackage.Data.ToByteArray();
        var hash = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        if (!hash.Equals(opaque.SourcePackage.Sha256, StringComparison.OrdinalIgnoreCase))
            throw new CodecException("source_package_hash_mismatch", "Source package bytes do not match the hash recorded in the public wire envelope.");
        if (!string.IsNullOrWhiteSpace(source?.PackageSha256) && !hash.Equals(source.PackageSha256, StringComparison.OrdinalIgnoreCase))
            throw new CodecException("source_identity_mismatch", "Source package bytes do not match source.package_sha256.");

        var actual = ValidateAndCollectOpaque(bytes, limits, profile, includeSourcePackage: false);
        AssertOpaqueGraphMatches(opaque, actual, "source_package_graph_mismatch");
        return bytes;
    }

    internal static void AssertOpaqueGraphMatches(
        OpaqueOpcGraph expected,
        OpaqueOpcGraph actual,
        string code,
        Func<OpaqueOpcRelationship, bool>? ignoreRelationship = null,
        Func<OpaqueOpcPart, bool>? ignorePart = null)
    {
        var expectedParts = expected.Parts
            .Where(item => ignorePart?.Invoke(item) != true)
            .Select(PartSignature)
            .OrderBy(item => item, StringComparer.Ordinal)
            .ToArray();
        var actualParts = actual.Parts
            .Where(item => ignorePart?.Invoke(item) != true)
            .Select(PartSignature)
            .OrderBy(item => item, StringComparer.Ordinal)
            .ToArray();
        var expectedRelationships = expected.PackageRelationships
            .Where(item => ignoreRelationship?.Invoke(item) != true)
            .Select(RelationshipSignature)
            .OrderBy(item => item, StringComparer.Ordinal)
            .ToArray();
        var actualRelationships = actual.PackageRelationships
            .Where(item => ignoreRelationship?.Invoke(item) != true)
            .Select(RelationshipSignature)
            .OrderBy(item => item, StringComparer.Ordinal)
            .ToArray();
        if (!expectedParts.SequenceEqual(actualParts, StringComparer.Ordinal) ||
            !expectedRelationships.SequenceEqual(actualRelationships, StringComparer.Ordinal))
        {
            var expectedPartPaths = expected.Parts.Where(item => ignorePart?.Invoke(item) != true).ToDictionary(item => item.Path, PartSignature, StringComparer.OrdinalIgnoreCase);
            var actualPartPaths = actual.Parts.Where(item => ignorePart?.Invoke(item) != true).ToDictionary(item => item.Path, PartSignature, StringComparer.OrdinalIgnoreCase);
            var changedParts = expectedPartPaths.Keys.Concat(actualPartPaths.Keys).Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(path => !expectedPartPaths.TryGetValue(path, out var left) || !actualPartPaths.TryGetValue(path, out var right) || left != right)
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .Take(8)
                .ToArray();
            var relationshipChanged = !expectedRelationships.SequenceEqual(actualRelationships, StringComparer.Ordinal);
            var detail = changedParts.Length > 0 ? $" Changed parts: {string.Join(", ", changedParts)}." : string.Empty;
            if (relationshipChanged) detail += " Relationship inventory changed.";
            throw new CodecException(code, $"Opaque OPC parts or relationships do not match the validated source package graph.{detail}");
        }
    }

    private static string PartSignature(OpaqueOpcPart part)
    {
        var hash = part.Sha256;
        if (!part.Data.IsEmpty)
        {
            var dataHash = Convert.ToHexString(SHA256.HashData(part.Data.Span)).ToLowerInvariant();
            if (!string.IsNullOrWhiteSpace(hash) && !hash.Equals(dataHash, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("opaque_part_hash_mismatch", $"Opaque part {part.Path} does not match its recorded hash.", part.Path);
            hash = dataHash;
        }
        return $"{part.Path}\0{hash.ToLowerInvariant()}";
    }

    private static string RelationshipSignature(OpaqueOpcRelationship relationship) =>
        $"{relationship.SourcePath}\0{relationship.Id}\0{relationship.Type}\0{relationship.Target}\0{relationship.TargetMode}";

    private static void CollectOpaqueRelationships(string relationshipPath, byte[] data, OpaqueOpcGraph opaque, OpcPackageProfile profile)
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
                if (!profile.OwnsRelationship(relationship)) opaque.PackageRelationships.Add(relationship);
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

    private static void ValidateEntryPath(string path, OpcPackageProfile profile)
    {
        if (string.IsNullOrWhiteSpace(path) || path.StartsWith("/", StringComparison.Ordinal) || path.Contains('\\'))
            throw new CodecException("unsafe_part_path", $"{profile.Format} package contains unsafe part path {path}.", path);
        var segments = path.Split('/', StringSplitOptions.RemoveEmptyEntries);
        if (segments.Any(segment => segment is "." or ".."))
            throw new CodecException("unsafe_part_path", $"{profile.Format} package contains unsafe part path {path}.", path);
    }
}
