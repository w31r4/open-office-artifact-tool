using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenChestnut.Codec;

// Projects read-only PresentationML objects onto the public artifact model.
// OpaqueOpcGraph remains the only owner of package metadata and source bytes;
// this catalog records only the relationship roots and reachable part paths.
internal sealed class PptxNativeObjectCatalog
{
    private const string SpreadsheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    private const string DocumentContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    private const string PowerPoint2010Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";
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

    private static readonly HashSet<string> DiagramNamespaces = new(StringComparer.Ordinal)
    {
        "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        "http://purl.oclc.org/ooxml/drawingml/diagram",
    };

    private static readonly HashSet<string> DrawingNamespaces = new(StringComparer.Ordinal)
    {
        "http://schemas.openxmlformats.org/drawingml/2006/main",
        "http://purl.oclc.org/ooxml/drawingml/main",
    };

    internal static bool IsRelationshipNamespace(string namespaceUri) =>
        RelationshipNamespaces.Contains(namespaceUri);

    internal static bool IsDiagramRelationshipIds(OpenXmlElement element) =>
        element.LocalName == "relIds" && DiagramNamespaces.Contains(element.NamespaceUri);

    // A video picture can otherwise satisfy the bounded poster-image reader
    // and be exposed as an ordinary editable image. Detect any native media
    // marker before semantic picture projection; the stricter clone preflight
    // later accepts only the exact closed MP4 graph.
    internal static bool IsMediaPicture(OpenXmlElement source)
    {
        if (source is not P.Picture) return false;
        return Elements(source).Any(element =>
            (element.LocalName is "videoFile" or "audioFile" && DrawingNamespaces.Contains(element.NamespaceUri)) ||
            (element.LocalName == "media" && element.NamespaceUri == PowerPoint2010Namespace));
    }

    private readonly EffectiveCodecLimits _limits;
    private readonly Dictionary<string, OpaqueOpcPart> _parts;
    private readonly Dictionary<string, OpaqueOpcRelationship> _relationships;
    private readonly HashSet<string> _packagePaths;
    private readonly Dictionary<string, string[]> _closureByTarget = new(StringComparer.OrdinalIgnoreCase);
    private ulong _referenceCount;
    private ulong _traversalCount;

    internal PptxNativeObjectCatalog(OpaqueOpcGraph opaque, byte[] sourceBytes, EffectiveCodecLimits limits)
    {
        _limits = limits;
        _parts = opaque.Parts.ToDictionary(part => part.Path, StringComparer.OrdinalIgnoreCase);
        _relationships = new Dictionary<string, OpaqueOpcRelationship>(StringComparer.OrdinalIgnoreCase);
        foreach (var relationship in opaque.PackageRelationships)
        {
            var key = RelationshipKey(relationship.SourcePath, relationship.Id);
            if (!_relationships.TryAdd(key, relationship))
                throw new CodecException(
                    "duplicate_relationship_id",
                    $"OPC part {relationship.SourcePath} contains duplicate relationship ID {relationship.Id}.",
                    relationship.SourcePath);
        }

        using var stream = new MemoryStream(sourceBytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        _packagePaths = archive.Entries
            .Where(entry => !entry.FullName.EndsWith("/", StringComparison.Ordinal))
            .Select(entry => entry.FullName)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var relationship in opaque.PackageRelationships)
        {
            if (relationship.TargetMode.Equals("External", StringComparison.OrdinalIgnoreCase)) continue;
            var targetPath = ResolveTarget(relationship.SourcePath, relationship.Target);
            if (!_packagePaths.Contains(targetPath))
                throw new CodecException(
                    "missing_presentation_native_part",
                    $"PPTX relationship {relationship.Id} from {relationship.SourcePath} references missing part {targetPath}.",
                    targetPath);
        }
    }

    internal void Populate(PresentationOpaqueElement target, OpenXmlElement source, OpenXmlPart owner)
    {
        var sourcePart = PartPath(owner);
        target.NativeKind = Classify(source);
        var seenReferences = new HashSet<string>(StringComparer.Ordinal);
        var rootRelationships = new List<OpaqueOpcRelationship>();
        var seenRelationshipIds = new HashSet<string>(StringComparer.Ordinal);

        foreach (var element in Elements(source))
        {
            foreach (var attribute in element.GetAttributes())
            {
                if (!RelationshipNamespaces.Contains(attribute.NamespaceUri)) continue;
                if (string.IsNullOrWhiteSpace(attribute.Value) && IsMediaActionSentinel(element, attribute)) continue;
                _referenceCount++;
                if (_referenceCount > _limits.MaxCells)
                    throw new CodecException(
                        "presentation_native_graph_budget_exceeded",
                        $"PPTX native-object relationship references exceed max_cells ({_limits.MaxCells}).",
                        sourcePart);
                if (string.IsNullOrWhiteSpace(attribute.Value))
                    throw new CodecException(
                        "missing_presentation_native_relationship",
                        $"PPTX native object in {sourcePart} contains an empty relationship reference.",
                        sourcePart);

                var attributeName = string.IsNullOrWhiteSpace(attribute.Prefix)
                    ? attribute.LocalName
                    : $"{attribute.Prefix}:{attribute.LocalName}";
                var referenceKey = $"{attributeName}\0{attribute.Value}";
                if (seenReferences.Add(referenceKey))
                {
                    target.RelationshipReferences.Add(new PresentationNativeRelationshipReference
                    {
                        Attribute = attributeName,
                        RelationshipId = attribute.Value,
                        NamespaceUri = attribute.NamespaceUri,
                    });
                }

                if (!seenRelationshipIds.Add(attribute.Value)) continue;
                if (!_relationships.TryGetValue(RelationshipKey(sourcePart, attribute.Value), out var relationship))
                    throw new CodecException(
                        "missing_presentation_native_relationship",
                        $"PPTX native object in {sourcePart} references missing relationship {attribute.Value}.",
                        sourcePart);
                rootRelationships.Add(relationship);
            }
        }

        var seenPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var relationship in rootRelationships)
        {
            if (relationship.TargetMode.Equals("External", StringComparison.OrdinalIgnoreCase)) continue;
            var targetPath = ResolveTarget(sourcePart, relationship.Target);
            foreach (var partPath in Closure(targetPath, sourcePart))
                if (seenPaths.Add(partPath)) target.PreservedPartPaths.Add(partPath);
        }

        TryPopulateOleWorkbook(target, source, sourcePart);
        TryPopulateOleOfficePackage(target, source, sourcePart);
        TryPopulateDiagramText(target, source, owner);
    }

    private void TryPopulateDiagramText(PresentationOpaqueElement target, OpenXmlElement source, OpenXmlPart owner)
    {
        if (target.NativeKind != "diagram" || !PptxDiagramTextCodec.TryDescribe(source, owner, out var diagram)) return;
        if (!_parts.TryGetValue(diagram.PartPath, out var part) ||
            !part.ContentType.Equals(diagram.ContentType, StringComparison.OrdinalIgnoreCase) ||
            !part.Sha256.Equals(diagram.SourceSha256, StringComparison.OrdinalIgnoreCase) ||
            !target.PreservedPartPaths.Contains(diagram.PartPath, StringComparer.OrdinalIgnoreCase))
            return;
        target.DiagramText = diagram;
    }

    private void TryPopulateOleWorkbook(PresentationOpaqueElement target, OpenXmlElement source, string sourcePart)
    {
        if (!TryResolveEditableOlePackage(target, source, sourcePart, out var part, out var relationshipId)) return;
        if (!part.ContentType.Equals(SpreadsheetContentType, StringComparison.OrdinalIgnoreCase)) return;

        target.OleWorkbook = new PresentationOleWorkbook
        {
            PartPath = part.Path,
            ContentType = SpreadsheetContentType,
            SourceSha256 = part.Sha256.ToLowerInvariant(),
            RelationshipId = relationshipId,
        };
    }

    private void TryPopulateOleOfficePackage(PresentationOpaqueElement target, OpenXmlElement source, string sourcePart)
    {
        if (!TryResolveEditableOlePackage(target, source, sourcePart, out var part, out var relationshipId)) return;
        if (!part.ContentType.Equals(DocumentContentType, StringComparison.OrdinalIgnoreCase)) return;
        target.OleOfficePackage = new PresentationOleOfficePackage
        {
            PartPath = part.Path,
            ContentType = DocumentContentType,
            SourceSha256 = part.Sha256.ToLowerInvariant(),
            RelationshipId = relationshipId,
            Kind = "docx",
        };
    }

    // Replacing a shared embedded package would affect more than the selected
    // OLE object. This helper exposes just the common, source-bound package
    // proof; concrete codecs decide which content type they are willing to
    // make editable.
    private bool TryResolveEditableOlePackage(
        PresentationOpaqueElement target,
        OpenXmlElement source,
        string sourcePart,
        out OpaqueOpcPart part,
        out string relationshipId)
    {
        part = null!;
        relationshipId = string.Empty;
        if (target.NativeKind != "oleObject" || source is not P.GraphicFrame || !SupportsPlacementEditing(source)) return false;
        var oleObjects = Elements(source)
            .Where(element => element.LocalName == "oleObj" && PresentationNamespaces.Contains(element.NamespaceUri))
            .ToArray();
        if (oleObjects.Length != 1) return false;
        var relationshipAttributes = oleObjects[0].GetAttributes()
            .Where(attribute => attribute.LocalName == "id" && RelationshipNamespaces.Contains(attribute.NamespaceUri))
            .ToArray();
        if (relationshipAttributes.Length != 1 || string.IsNullOrWhiteSpace(relationshipAttributes[0].Value)) return false;
        relationshipId = relationshipAttributes[0].Value!;
        if (!_relationships.TryGetValue(RelationshipKey(sourcePart, relationshipId), out var relationship) ||
            relationship.TargetMode.Equals("External", StringComparison.OrdinalIgnoreCase) ||
            !relationship.Type.EndsWith("/package", StringComparison.Ordinal)) return false;
        var targetPath = ResolveTarget(sourcePart, relationship.Target);
        if (!_parts.TryGetValue(targetPath, out var resolvedPart) ||
            resolvedPart.Sha256.Length != 64 || !resolvedPart.Sha256.All(char.IsAsciiHexDigit) ||
            !target.PreservedPartPaths.Contains(targetPath, StringComparer.OrdinalIgnoreCase)) return false;
        part = resolvedPart;
        var inboundCount = _relationships.Values.Count(candidate =>
            !candidate.TargetMode.Equals("External", StringComparison.OrdinalIgnoreCase) &&
            ResolveTarget(candidate.SourcePath, candidate.Target).Equals(targetPath, StringComparison.OrdinalIgnoreCase));
        return inboundCount == 1;
    }

    // Placement editing is intentionally narrower than native-object graph
    // discovery. Only the three recognized top-level roots with complete
    // owner transforms and non-visual names are editable; every relationship,
    // part, descendant and other attribute remains source-bound.
    internal static bool SupportsPlacementEditing(OpenXmlElement source)
    {
        var kind = Classify(source);
        if (kind is "oleObject" or "diagram" && source is P.GraphicFrame frame)
        {
            return frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties is not null &&
                   frame.Transform?.Offset is not null &&
                   frame.Transform.Extents is not null;
        }
        if (kind == "contentPart" && source is P.GroupShape group)
        {
            var transform = group.GetFirstChild<P.GroupShapeProperties>()?.GetFirstChild<A.TransformGroup>();
            return group.GetFirstChild<P.NonVisualGroupShapeProperties>()?.NonVisualDrawingProperties is not null &&
                   transform?.Offset is not null &&
                   transform.Extents is not null &&
                   transform.ChildOffset is not null &&
                   transform.ChildExtents is not null;
        }
        return false;
    }

    private string[] Closure(string rootPath, string sourcePart)
    {
        if (_closureByTarget.TryGetValue(rootPath, out var cached)) return cached;
        if (!_packagePaths.Contains(rootPath))
            throw new CodecException(
                "missing_presentation_native_part",
                $"PPTX native object in {sourcePart} references missing part {rootPath}.",
                rootPath);

        var ordered = new List<string>();
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var pending = new Queue<string>();
        pending.Enqueue(rootPath);
        while (pending.Count > 0)
        {
            var path = pending.Dequeue();
            if (!visited.Add(path)) continue;
            _traversalCount++;
            if (_traversalCount > _limits.MaxCells)
                throw new CodecException(
                    "presentation_native_graph_budget_exceeded",
                    $"PPTX native-object graph traversal exceeds max_cells ({_limits.MaxCells}).",
                    sourcePart);
            if (!_packagePaths.Contains(path))
                throw new CodecException(
                    "missing_presentation_native_part",
                    $"PPTX native object in {sourcePart} references missing part {path}.",
                    path);

            // Modeled PresentationML parts are intentionally absent from the
            // opaque inventory. Their existence is still checked above, but
            // the native graph must not claim their semantic subgraph.
            if (!_parts.TryGetValue(path, out var part)) continue;
            ordered.Add(path);
            if ((uint)ordered.Count > _limits.MaxParts)
                throw new CodecException(
                    "presentation_native_graph_budget_exceeded",
                    $"PPTX native object reaches more than max_parts ({_limits.MaxParts}) package parts.",
                    sourcePart);
            foreach (var relationship in part.Relationships)
            {
                if (relationship.TargetMode.Equals("External", StringComparison.OrdinalIgnoreCase)) continue;
                pending.Enqueue(ResolveTarget(path, relationship.Target));
            }
        }

        var result = ordered.ToArray();
        _closureByTarget[rootPath] = result;
        return result;
    }

    internal static string Classify(OpenXmlElement source)
    {
        var descendants = Elements(source).ToArray();
        if (source is P.Picture && IsMediaPicture(source))
            return "media";
        if (descendants.Any(element =>
                element.LocalName == "oleObj" && PresentationNamespaces.Contains(element.NamespaceUri)))
            return "oleObject";
        if (descendants.Any(element =>
                (element.LocalName == "relIds" && DiagramNamespaces.Contains(element.NamespaceUri)) ||
                (element.LocalName == "graphicData" && element.GetAttributes().Any(attribute =>
                    attribute.LocalName == "uri" && attribute.Value is { } value && DiagramNamespaces.Contains(value)))))
            return "diagram";
        if (descendants.Any(element =>
                element.LocalName == "contentPart" && PresentationNamespaces.Contains(element.NamespaceUri)))
            return "contentPart";
        return source.LocalName switch
        {
            "pic" => "picture",
            "grpSp" => "group",
            "cxnSp" => "connector",
            "graphicFrame" => "graphicFrame",
            _ => string.IsNullOrWhiteSpace(source.LocalName) ? "nativeObject" : source.LocalName,
        };
    }

    private static bool IsMediaActionSentinel(OpenXmlElement element, OpenXmlAttribute attribute) =>
        element.LocalName == "hlinkClick" &&
        DrawingNamespaces.Contains(element.NamespaceUri) &&
        attribute.LocalName == "id" &&
        RelationshipNamespaces.Contains(attribute.NamespaceUri) &&
        element.GetAttributes().Any(candidate =>
            candidate.LocalName == "action" &&
            candidate.NamespaceUri.Length == 0 &&
            candidate.Value == "ppaction://media");

    // Keep the OPC target resolver shared with the narrow source-preserving
    // presentation topology operations. The source package has already passed
    // the package guard before this is used, but topology edits still need the
    // same dot-segment and traversal semantics when they inspect inbound links.
    internal static string ResolveTarget(string sourcePath, string target)
    {
        if (string.IsNullOrWhiteSpace(target) || target.Any(char.IsControl) || target.Contains('\\') || target.Contains('?'))
            throw InvalidTarget(sourcePath, target);
        var partTarget = target.Split('#', 2)[0];
        if (partTarget.Length == 0 || partTarget.Contains(':')) throw InvalidTarget(sourcePath, target);

        var segments = new List<string>();
        if (!partTarget.StartsWith("/", StringComparison.Ordinal))
        {
            var separator = sourcePath.LastIndexOf('/');
            if (separator >= 0) segments.AddRange(sourcePath[..separator].Split('/', StringSplitOptions.RemoveEmptyEntries));
        }
        foreach (var segment in partTarget.TrimStart('/').Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (segment == ".") continue;
            if (segment == "..")
            {
                if (segments.Count == 0) throw InvalidTarget(sourcePath, target);
                segments.RemoveAt(segments.Count - 1);
                continue;
            }
            segments.Add(segment);
        }
        if (segments.Count == 0) throw InvalidTarget(sourcePath, target);
        return string.Join('/', segments);
    }

    private static CodecException InvalidTarget(string sourcePath, string target) => new(
        "invalid_presentation_native_target",
        $"PPTX native object relationship from {sourcePath} has unsafe target {target}.",
        sourcePath);

    private static string RelationshipKey(string sourcePath, string relationshipId) => $"{sourcePath}\0{relationshipId}";

    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');

    private static IEnumerable<OpenXmlElement> Elements(OpenXmlElement source)
    {
        yield return source;
        foreach (var descendant in source.Descendants()) yield return descendant;
    }
}
