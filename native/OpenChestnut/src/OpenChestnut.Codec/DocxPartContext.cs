using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Owns relationships whose source is word/document.xml. Relationship IDs are
// package-local locators; edits track only the IDs that were added or removed
// so the generic opaque-package guard can continue protecting every other
// relationship, including hyperlinks inside unsupported structures.
internal sealed class DocxPartContext
{
    private readonly HashSet<string> _mutatedRelationshipIds = new(StringComparer.Ordinal);
    private readonly HashSet<string> _mutatedPartPaths = new(StringComparer.OrdinalIgnoreCase);
    private XDocument? _numberingDocument;
    private XDocument? _stylesDocument;
    private bool _numberingDocumentLoaded;
    private string? _mutatedNumberingPartPath;
    private string? _mutatedCommentsPartPath;
    private string? _mutatedCommentsRelationshipId;
    private bool _stylesDocumentLoaded;
    private uint? _nextDrawingId;
    private readonly HashSet<string> _plannedBookmarks;

    internal DocxPartContext(MainDocumentPart owner, DocxImageAssetCatalog? images = null, IEnumerable<string>? plannedBookmarks = null)
    {
        Owner = owner;
        Images = images;
        _plannedBookmarks = new HashSet<string>(plannedBookmarks ?? [], StringComparer.Ordinal);
    }

    internal MainDocumentPart Owner { get; }
    internal DocxImageAssetCatalog? Images { get; }
    internal IReadOnlyCollection<string> MutatedRelationshipIds => _mutatedRelationshipIds;

    internal uint NextDrawingId()
    {
        _nextDrawingId ??= Math.Max(1U, Owner.Document?.Descendants<W.Drawing>()
            .SelectMany(drawing => drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>())
            .Select(properties => properties.Id?.Value ?? 0U)
            .DefaultIfEmpty(0U)
            .Max() + 1U ?? 1U);
        var result = _nextDrawingId.Value;
        _nextDrawingId = result + 1U;
        return result;
    }

    internal void MarkPartMutated(OpenXmlPart part) =>
        _mutatedPartPaths.Add(part.Uri.OriginalString.TrimStart('/'));

    // Read semantic support parts without materializing an Open XML SDK root.
    // This prevents AutoSave from normalizing untouched source XML during a
    // document.xml-only edit and amortizes parsing across all body blocks.
    internal XDocument? NumberingDocument => ReadCachedPart(
        Owner.NumberingDefinitionsPart,
        ref _numberingDocumentLoaded,
        ref _numberingDocument);

    internal XDocument? StylesDocument => ReadCachedPart(
        Owner.StyleDefinitionsPart,
        ref _stylesDocumentLoaded,
        ref _stylesDocument);

    internal void SaveNumberingDocument()
    {
        var document = NumberingDocument ?? throw new CodecException(
            "unsupported_document_edit",
            "Numbering-definition edits require a source Numbering part.",
            "word/numbering.xml");
        var part = Owner.NumberingDefinitionsPart ?? throw new CodecException(
            "unsupported_document_edit",
            "Numbering-definition edits require a source Numbering part.",
            "word/numbering.xml");
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = XmlWriter.Create(stream, new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            OmitXmlDeclaration = document.Declaration is null,
            Indent = false,
        });
        document.Save(writer);
        _mutatedNumberingPartPath = part.Uri.OriginalString.TrimStart('/');
    }

    internal bool HasNumberingReferenceOutsideMainDocument(int numberingId)
    {
        var visited = new HashSet<OpenXmlPart>(ReferenceEqualityComparer.Instance);
        var pending = new Stack<OpenXmlPart>(Owner.Parts.Select(pair => pair.OpenXmlPart));
        while (pending.Count > 0)
        {
            var part = pending.Pop();
            if (!visited.Add(part) || ReferenceEquals(part, Owner.NumberingDefinitionsPart)) continue;
            foreach (var child in part.Parts) pending.Push(child.OpenXmlPart);
            if (!part.ContentType.Contains("xml", StringComparison.OrdinalIgnoreCase)) continue;
            try
            {
                using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
                using var reader = XmlReader.Create(stream, new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                });
                var document = XDocument.Load(reader, LoadOptions.None);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                if (document.Descendants(w + "numId").Any(element =>
                        int.TryParse(element.Attribute(w + "val")?.Value, out var value) && value == numberingId))
                    return true;
            }
            catch (XmlException)
            {
                return true;
            }
        }
        return false;
    }

    internal bool HasBookmark(string name) =>
        _plannedBookmarks.Contains(name) ||
        Owner.Document?.Descendants<W.BookmarkStart>().Any(item => item.Name?.Value == name) == true;

    internal bool TryReadExternal(string relationshipId, out string uri)
    {
        uri = string.Empty;
        if (string.IsNullOrWhiteSpace(relationshipId)) return false;
        var relationship = Owner.HyperlinkRelationships.FirstOrDefault(item => item.Id == relationshipId);
        if (relationship is null || !relationship.IsExternal) return false;
        uri = relationship.Uri.OriginalString;
        return true;
    }

    internal string EnsureExternal(string uri, string currentRelationshipId)
    {
        if (TryReadExternal(currentRelationshipId, out var currentUri) && currentUri.Equals(uri, StringComparison.Ordinal))
            return currentRelationshipId;

        var existing = Owner.HyperlinkRelationships.FirstOrDefault(item =>
            item.IsExternal && item.Uri.OriginalString.Equals(uri, StringComparison.Ordinal));
        if (existing is not null) return existing.Id;

        var relationship = Owner.AddHyperlinkRelationship(new Uri(uri, UriKind.Absolute), true);
        _mutatedRelationshipIds.Add(relationship.Id);
        return relationship.Id;
    }

    internal void RemoveIfUnreferenced(string relationshipId)
    {
        if (string.IsNullOrWhiteSpace(relationshipId)) return;
        if (Owner.Document?.Descendants().Any(element => element.GetAttributes().Any(attribute =>
                attribute.LocalName == "id" &&
                attribute.Value == relationshipId &&
                attribute.NamespaceUri.EndsWith("/relationships", StringComparison.Ordinal))) == true) return;
        if (Owner.HyperlinkRelationships.All(item => item.Id != relationshipId)) return;
        Owner.DeleteReferenceRelationship(relationshipId);
        _mutatedRelationshipIds.Add(relationshipId);
    }

    internal void MarkCommentsMutated(WordprocessingCommentsPart part)
    {
        var pair = Owner.Parts.FirstOrDefault(item => ReferenceEquals(item.OpenXmlPart, part));
        if (pair.OpenXmlPart is null)
            throw new CodecException(
                "document_comment_source_binding_mismatch",
                "The modeled Comments part is not related from word/document.xml.",
                part.Uri.OriginalString.TrimStart('/'));
        _mutatedCommentsRelationshipId = pair.RelationshipId;
        _mutatedCommentsPartPath = part.Uri.OriginalString.TrimStart('/');
    }

    internal bool IgnoresModeledRelationship(OpenOffice.Artifact.Wire.V1.OpaqueOpcRelationship relationship) =>
        relationship.SourcePath.Equals("word/document.xml", StringComparison.OrdinalIgnoreCase) &&
        ((relationship.Type.EndsWith("/hyperlink", StringComparison.Ordinal) &&
          _mutatedRelationshipIds.Contains(relationship.Id)) ||
         (relationship.Type.EndsWith("/comments", StringComparison.Ordinal) &&
          relationship.Id.Equals(_mutatedCommentsRelationshipId, StringComparison.Ordinal)));

    internal bool IgnoresModeledPart(OpenOffice.Artifact.Wire.V1.OpaqueOpcPart part) =>
        _mutatedPartPaths.Contains(part.Path) ||
        (_mutatedNumberingPartPath is not null &&
         part.Path.Equals(_mutatedNumberingPartPath, StringComparison.OrdinalIgnoreCase)) ||
        (_mutatedCommentsPartPath is not null &&
         part.Path.Equals(_mutatedCommentsPartPath, StringComparison.OrdinalIgnoreCase));

    private static XDocument? ReadCachedPart(OpenXmlPart? part, ref bool loaded, ref XDocument? document)
    {
        if (loaded) return document;
        loaded = true;
        if (part is null) return null;
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
        });
        document = XDocument.Load(reader, LoadOptions.None);
        return document;
    }
}
