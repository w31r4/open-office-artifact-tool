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

    internal DocxPartContext(MainDocumentPart owner)
    {
        Owner = owner;
    }

    internal MainDocumentPart Owner { get; }
    internal IReadOnlyCollection<string> MutatedRelationshipIds => _mutatedRelationshipIds;

    internal bool HasBookmark(string name) =>
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

    internal bool IgnoresModeledRelationship(OpenOffice.Artifact.Wire.V1.OpaqueOpcRelationship relationship) =>
        relationship.SourcePath.Equals("word/document.xml", StringComparison.OrdinalIgnoreCase) &&
        relationship.Type.EndsWith("/hyperlink", StringComparison.Ordinal) &&
        _mutatedRelationshipIds.Contains(relationship.Id);
}
