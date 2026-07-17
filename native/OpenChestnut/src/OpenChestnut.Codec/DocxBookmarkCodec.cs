using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Owns one deliberately reversible bookmark profile: a matching w:bookmarkStart
// and w:bookmarkEnd wrap all content in one paragraph-like body block. More
// general nested, crossing, table-cell, or multi-block ranges stay source-bound.
internal static class DocxBookmarkCodec
{
    private sealed record CanonicalBookmark(string Name, string NativeId, uint BodyIndex, string BlockId);
    private sealed record BookmarkIdentity(
        string Description,
        string Name,
        string NativeId,
        Action<string> AssignNativeId);

    internal static IEnumerable<string> Names(DocumentArtifact document) =>
        Identities(document).Select(identity => identity.Name);

    internal static void AssignNativeIds(DocumentArtifact document)
    {
        var identities = Identities(document).ToArray();
        ValidateIdentitySpace(identities, requireNativeIds: false);
        var usedNativeIds = identities
            .Where(identity => identity.NativeId.Length > 0)
            .Select(identity => identity.NativeId)
            .ToHashSet(StringComparer.Ordinal);
        uint nextNativeId = 0;
        foreach (var identity in identities.Where(identity => identity.NativeId.Length == 0))
        {
            while (usedNativeIds.Contains(nextNativeId.ToString())) nextNativeId++;
            var nativeId = nextNativeId.ToString();
            identity.AssignNativeId(nativeId);
            usedNativeIds.Add(nativeId);
            nextNativeId++;
        }
    }

    internal static void Read(
        W.Body body,
        DocumentArtifact document,
        ref ulong semanticItems,
        EffectiveCodecLimits limits)
    {
        var blockByBodyIndex = document.Blocks
            .Where(block => block.Source is not null)
            .ToDictionary(block => block.Source.BodyIndex);
        foreach (var source in ReadCanonical(body, blockByBodyIndex))
        {
            var bookmark = new DocumentBookmark
            {
                Id = $"document/bookmark/{document.Bookmarks.Count + 1}",
                Name = source.Name,
                TargetBlockId = source.BlockId,
                EndTargetBlockId = source.BlockId,
                NativeId = source.NativeId,
                Source = new DocumentBookmarkSourceBinding
                {
                    BodyIndex = source.BodyIndex,
                    NativeId = source.NativeId,
                    Editable = false,
                },
            };
            bookmark.Source.SemanticSha256 = SemanticHash(bookmark);
            document.Bookmarks.Add(bookmark);
            semanticItems++;
            if (semanticItems > limits.MaxCells)
                throw new CodecException(
                    "document_item_budget_exceeded",
                    $"DOCX document exceeds max_cells semantic-item budget ({limits.MaxCells}).",
                    "word/document.xml");
        }
    }

    internal static void Author(W.Body body, DocumentArtifact document)
    {
        Validate(document);
        var blockIndexes = document.Blocks
            .Select((block, index) => (block.Id, Index: index))
            .ToDictionary(item => item.Id, item => item.Index, StringComparer.Ordinal);
        foreach (var bookmark in document.Bookmarks)
        {
            var nativeId = bookmark.NativeId;
            if (string.IsNullOrEmpty(nativeId)) throw Invalid($"Document bookmark {bookmark.Id} is missing its planned native ID.");

            var element = body.ChildElements[blockIndexes[bookmark.TargetBlockId]];
            if (element is not W.Paragraph paragraph)
                throw Invalid($"Document bookmark {bookmark.Id} target must serialize as a paragraph.");
            var start = new W.BookmarkStart { Id = nativeId, Name = bookmark.Name };
            var end = new W.BookmarkEnd { Id = nativeId };
            var firstContent = paragraph.ChildElements.FirstOrDefault(child => child is not W.ParagraphProperties);
            if (firstContent is null) paragraph.Append(start, end);
            else
            {
                paragraph.InsertBefore(start, firstContent);
                paragraph.Append(end);
            }
        }
    }

    internal static void AssertSourceUnchanged(byte[] sourceBytes, DocumentArtifact requested)
    {
        using var stream = new MemoryStream(sourceBytes, writable: false);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var body = package.MainDocumentPart?.Document?.Body ??
            throw new CodecException("missing_document_body", "DOCX package has no document body.", "word/document.xml");
        var blockByBodyIndex = requested.Blocks
            .Where(block => block.Source is not null)
            .ToDictionary(block => block.Source.BodyIndex);
        var source = ReadCanonical(body, blockByBodyIndex).ToArray();
        if (source.Length != requested.Bookmarks.Count)
            throw new CodecException(
                "document_bookmark_topology_changed",
                $"Source-preserving DOCX export requires the original {source.Length}-bookmark bounded topology; the artifact contains {requested.Bookmarks.Count} bookmarks.",
                "word/document.xml");
        for (var index = 0; index < source.Length; index++)
        {
            var actual = source[index];
            var bookmark = requested.Bookmarks[index];
            var binding = bookmark.Source ?? throw new CodecException(
                "missing_document_bookmark_source_binding",
                $"Imported document bookmark {bookmark.Id} is missing its source binding.",
                "word/document.xml");
            if (binding.BodyIndex != actual.BodyIndex ||
                !binding.NativeId.Equals(actual.NativeId, StringComparison.Ordinal) ||
                !bookmark.NativeId.Equals(actual.NativeId, StringComparison.Ordinal) ||
                !bookmark.Name.Equals(actual.Name, StringComparison.Ordinal) ||
                !bookmark.TargetBlockId.Equals(actual.BlockId, StringComparison.Ordinal) ||
                !bookmark.EndTargetBlockId.Equals(actual.BlockId, StringComparison.Ordinal) ||
                !binding.SemanticSha256.Equals(SemanticHash(bookmark), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "unsupported_document_bookmark_edit",
                    $"Imported document bookmark {bookmark.Id} identity, name, and target are source-bound in protocol 2.",
                    "word/document.xml");
        }
    }

    internal static void Validate(DocumentArtifact document)
    {
        ValidateIdentitySpace(Identities(document).ToArray(), requireNativeIds: true);
        var blocks = document.Blocks.ToDictionary(block => block.Id, StringComparer.Ordinal);
        var targets = new HashSet<string>(StringComparer.Ordinal);
        foreach (var bookmark in document.Bookmarks)
        {
            if (string.IsNullOrWhiteSpace(bookmark.TargetBlockId) ||
                !bookmark.TargetBlockId.Equals(bookmark.EndTargetBlockId, StringComparison.Ordinal))
                throw Invalid($"Document bookmark {bookmark.Id} must wrap exactly one block in protocol 2.");
            if (!targets.Add(bookmark.TargetBlockId))
                throw Invalid($"Document bookmark target {bookmark.TargetBlockId} already has a bounded bookmark.");
            if (!blocks.TryGetValue(bookmark.TargetBlockId, out var block) || !IsParagraphLike(block))
                throw Invalid($"Document bookmark {bookmark.Id} target must be a paragraph, hyperlink, field, citation, tracked change, or image block.");
            if (block.ContentCase == DocumentBlock.ContentOneofCase.Paragraph &&
                block.Paragraph.Runs.Any(run => run.InlineField?.BookmarkName.Length > 0))
                throw Invalid($"Document bookmark {bookmark.Id} cannot wrap a paragraph that already contains a bounded inline-field bookmark.");
            if (bookmark.Source is not null && bookmark.Source.Editable)
                throw Invalid($"Imported document bookmark {bookmark.Id} cannot claim editable source topology in protocol 2.");
        }
    }

    private static IEnumerable<BookmarkIdentity> Identities(DocumentArtifact document)
    {
        foreach (var bookmark in document.Bookmarks)
        {
            var captured = bookmark;
            yield return new BookmarkIdentity(
                $"Document bookmark {captured.Id}",
                captured.Name,
                captured.NativeId,
                value => captured.NativeId = value);
        }
        foreach (var block in document.Blocks.Where(block => block.ContentCase == DocumentBlock.ContentOneofCase.Paragraph))
        {
            for (var runIndex = 0; runIndex < block.Paragraph.Runs.Count; runIndex++)
            {
                var field = block.Paragraph.Runs[runIndex].InlineField;
                if (field is null || field.BookmarkName.Length == 0 && field.BookmarkNativeId.Length == 0) continue;
                var captured = field;
                yield return new BookmarkIdentity(
                    $"Document paragraph {block.Id} inline field {runIndex}",
                    captured.BookmarkName,
                    captured.BookmarkNativeId,
                    value => captured.BookmarkNativeId = value);
            }
        }
    }

    private static void ValidateIdentitySpace(IReadOnlyList<BookmarkIdentity> identities, bool requireNativeIds)
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var nativeIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var identity in identities)
        {
            if (!ValidName(identity.Name))
                throw Invalid($"{identity.Description} name must start with an ASCII letter and contain only letters, digits, or underscores (maximum 40 characters).");
            if (!names.Add(identity.Name)) throw Invalid($"Document bookmark name {identity.Name} is duplicated.");
            if (identity.NativeId.Length == 0)
            {
                if (requireNativeIds) throw Invalid($"{identity.Description} is missing its planned native ID.");
                continue;
            }
            if (!uint.TryParse(identity.NativeId, out _))
                throw Invalid($"{identity.Description} native ID must be an unsigned decimal integer when present.");
            if (!nativeIds.Add(identity.NativeId))
                throw Invalid($"{identity.Description} duplicates native ID {identity.NativeId}.");
        }
    }

    private static IEnumerable<CanonicalBookmark> ReadCanonical(
        W.Body body,
        IReadOnlyDictionary<uint, DocumentBlock> blockByBodyIndex)
    {
        for (var bodyIndex = 0; bodyIndex < body.ChildElements.Count; bodyIndex++)
        {
            if (body.ChildElements[bodyIndex] is not W.Paragraph paragraph ||
                !blockByBodyIndex.TryGetValue(checked((uint)bodyIndex), out var block)) continue;
            var starts = paragraph.Elements<W.BookmarkStart>().ToArray();
            var ends = paragraph.Elements<W.BookmarkEnd>().ToArray();
            if (starts.Length != 1 || ends.Length != 1 ||
                paragraph.Descendants<W.BookmarkStart>().Count() != 1 ||
                paragraph.Descendants<W.BookmarkEnd>().Count() != 1) continue;
            var start = starts[0];
            var end = ends[0];
            var nativeId = start.Id?.Value ?? string.Empty;
            var name = start.Name?.Value ?? string.Empty;
            if (!ValidName(name) || !uint.TryParse(nativeId, out _) || end.Id?.Value != nativeId) continue;
            var children = paragraph.ChildElements.ToArray();
            var firstContent = Array.FindIndex(children, child => child is not W.ParagraphProperties);
            if (firstContent < 0 || !ReferenceEquals(children[firstContent], start) || !ReferenceEquals(children[^1], end)) continue;
            yield return new CanonicalBookmark(name, nativeId, checked((uint)bodyIndex), block.Id);
        }
    }

    private static bool IsParagraphLike(DocumentBlock block) => block.ContentCase is
        DocumentBlock.ContentOneofCase.Paragraph or
        DocumentBlock.ContentOneofCase.Hyperlink or
        DocumentBlock.ContentOneofCase.Field or
        DocumentBlock.ContentOneofCase.Citation or
        DocumentBlock.ContentOneofCase.Change or
        DocumentBlock.ContentOneofCase.Image;

    internal static bool ValidName(string value) =>
        value.Length is >= 1 and <= 40 &&
        ((value[0] is >= 'A' and <= 'Z') || (value[0] is >= 'a' and <= 'z')) &&
        value.All(character =>
            (character is >= 'A' and <= 'Z') ||
            (character is >= 'a' and <= 'z') ||
            (character is >= '0' and <= '9') || character == '_');

    private static string SemanticHash(DocumentBookmark bookmark)
    {
        var semantic = bookmark.Clone();
        semantic.Id = string.Empty;
        semantic.NativeId = string.Empty;
        semantic.Source = null;
        return Convert.ToHexString(SHA256.HashData(semantic.ToByteArray())).ToLowerInvariant();
    }

    private static CodecException Invalid(string message) => new("invalid_document_bookmark", message, "word/document.xml");
}
