using System.Buffers.Binary;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using W16Cid = DocumentFormat.OpenXml.Office2019.Word.Cid;
using W16Cex = DocumentFormat.OpenXml.Office2021.Word.CommentsExt;

namespace OpenChestnut.Codec;

internal sealed record DocxExtendedCommentInfo(
    string ParagraphId,
    string ParentParagraphId,
    bool Resolved,
    string DurableId,
    string DateUtc,
    DocumentCommentPerson? Person,
    bool IntelligentPlaceholder);

internal sealed record DocxExtendedCommentGraph(
    bool IsModern,
    IReadOnlyDictionary<string, DocxExtendedCommentInfo> ByNativeCommentId,
    WordprocessingCommentsExPart? CommentsExPart,
    WordprocessingCommentsIdsPart? CommentsIdsPart,
    WordCommentsExtensiblePart? CommentsExtensiblePart,
    WordprocessingPeoplePart? PeoplePart,
    string ExtendedGraphSha256,
    string CommentsIdsGraphSha256,
    string CommentsExtensibleGraphSha256,
    string PeopleGraphSha256);

// Owns the optional Word 2013+ support parts around classic w:comment bodies.
// The bounded profile is deliberately complete: every modern comment has one
// w14:paraId and one w15:commentEx; replies are direct children of roots;
// commentsIds/commentsExtensible/people, when present, are exact closed maps.
internal static class DocxExtendedCommentCodec
{
    private const string W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    private const string W15Ns = "http://schemas.microsoft.com/office/word/2012/wordml";
    private const string W16CidNs = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
    private const string W16CexNs = "http://schemas.microsoft.com/office/word/2018/wordml/cex";

    internal static bool TryRead(
        MainDocumentPart owner,
        IReadOnlyList<W.Comment> comments,
        out DocxExtendedCommentGraph graph,
        out string reason)
    {
        graph = null!;
        reason = string.Empty;
        var commentsExPart = owner.WordprocessingCommentsExPart;
        var commentsIdsPart = owner.WordprocessingCommentsIdsPart;
        var commentsExtensiblePart = owner.WordCommentsExtensiblePart;
        var peoplePart = owner.WordprocessingPeoplePart;
        var hasAny = commentsExPart is not null || commentsIdsPart is not null ||
                     commentsExtensiblePart is not null || peoplePart is not null;
        if (!hasAny)
        {
            graph = new DocxExtendedCommentGraph(
                false,
                new Dictionary<string, DocxExtendedCommentInfo>(StringComparer.Ordinal),
                null, null, null, null,
                string.Empty, string.Empty, string.Empty, string.Empty);
            return true;
        }
        if (commentsExPart?.CommentsEx is null)
        {
            reason = "modern comment support parts require one commentsExtended root";
            return false;
        }
        foreach (var part in new OpenXmlPart?[] { commentsExPart, commentsIdsPart, commentsExtensiblePart, peoplePart })
        {
            if (part is null) continue;
            if (part.Parts.Any() || part.ExternalRelationships.Any())
            {
                reason = $"comment support part {part.Uri} has a connected relationship graph";
                return false;
            }
        }

        var commentByParagraphId = new Dictionary<string, (string NativeId, W.Comment Element)>(StringComparer.Ordinal);
        foreach (var comment in comments)
        {
            var nativeId = comment.Id?.Value ?? string.Empty;
            var paragraph = comment.Elements<W.Paragraph>().SingleOrDefault();
            var paragraphId = NormalizeParagraphId(Attribute(paragraph, "paraId", W14));
            if (paragraph is null || paragraphId is null || !commentByParagraphId.TryAdd(paragraphId, (nativeId, comment)))
            {
                reason = $"modern comment {nativeId} requires one unique eight-digit w14:paraId";
                return false;
            }
        }

        var commentsEx = commentsExPart.CommentsEx;
        if (commentsEx.ChildElements.Any(child => child is not W15.CommentEx))
        {
            reason = "commentsExtended contains an unsupported child";
            return false;
        }
        var extendedByParagraphId = new Dictionary<string, (W15.CommentEx Element, string Parent, bool Done)>(StringComparer.Ordinal);
        foreach (var element in commentsEx.Elements<W15.CommentEx>())
        {
            var paragraphId = NormalizeParagraphId(Attribute(element, "paraId", W15Ns));
            var parent = NormalizeOptionalParagraphId(Attribute(element, "paraIdParent", W15Ns), out var parentValid);
            if (paragraphId is null || !parentValid || !TryBoolean(Attribute(element, "done", W15Ns), out var done) ||
                !extendedByParagraphId.TryAdd(paragraphId, (element, parent, done)))
            {
                reason = "commentsExtended requires unique paragraph IDs, valid direct-parent IDs, and boolean done values";
                return false;
            }
        }
        if (extendedByParagraphId.Count != comments.Count ||
            extendedByParagraphId.Keys.Any(id => !commentByParagraphId.ContainsKey(id)))
        {
            reason = "commentsExtended must map every classic comment exactly once";
            return false;
        }
        foreach (var (paragraphId, entry) in extendedByParagraphId)
        {
            if (entry.Parent.Length == 0) continue;
            if (!extendedByParagraphId.TryGetValue(entry.Parent, out var parent) || parent.Parent.Length != 0 || entry.Parent == paragraphId)
            {
                reason = $"comment paragraph {paragraphId} is not one direct reply to a root";
                return false;
            }
        }

        var durableByParagraphId = new Dictionary<string, string>(StringComparer.Ordinal);
        if (commentsIdsPart is not null)
        {
            var root = commentsIdsPart.CommentsIds;
            if (root is null || root.ChildElements.Any(child => child is not W16Cid.CommentId))
            {
                reason = "commentsIds has no supported root or contains an unsupported child";
                return false;
            }
            foreach (var element in root.Elements<W16Cid.CommentId>())
            {
                var paragraphId = NormalizeParagraphId(Attribute(element, "paraId", W16CidNs));
                var durableId = NormalizeDurableId(Attribute(element, "durableId", W16CidNs));
                if (paragraphId is null || durableId is null || !durableByParagraphId.TryAdd(paragraphId, durableId))
                {
                    reason = "commentsIds requires unique eight-digit paragraph/durable ID pairs";
                    return false;
                }
            }
            if (durableByParagraphId.Count != comments.Count || durableByParagraphId.Keys.Any(id => !commentByParagraphId.ContainsKey(id)) ||
                durableByParagraphId.Values.Distinct(StringComparer.Ordinal).Count() != durableByParagraphId.Count)
            {
                reason = "commentsIds must map every comment to one unique durable ID";
                return false;
            }
        }

        var extensibleByDurableId = new Dictionary<string, (string DateUtc, bool Intelligent)>(StringComparer.Ordinal);
        if (commentsExtensiblePart is not null)
        {
            if (commentsIdsPart is null)
            {
                reason = "commentsExtensible requires commentsIds";
                return false;
            }
            var root = commentsExtensiblePart.CommentsExtensible;
            if (root is null || root.ChildElements.Any(child => child is not W16Cex.CommentExtensible))
            {
                reason = "commentsExtensible has no supported root or contains an unsupported child";
                return false;
            }
            foreach (var element in root.Elements<W16Cex.CommentExtensible>())
            {
                if (element.ChildElements.Count != 0)
                {
                    reason = "commentsExtensible extension lists are outside the bounded profile";
                    return false;
                }
                var durableId = NormalizeDurableId(Attribute(element, "durableId", W16CexNs));
                var dateUtc = Attribute(element, "dateUtc", W16CexNs) ?? string.Empty;
                if (dateUtc.Length > 0 && !DateTimeOffset.TryParse(dateUtc, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out _))
                {
                    reason = $"comment durable ID {durableId} has an invalid UTC timestamp";
                    return false;
                }
                if (!TryBoolean(Attribute(element, "intelligentPlaceholder", W16CexNs), out var intelligent) ||
                    durableId is null || !extensibleByDurableId.TryAdd(durableId, (dateUtc, intelligent)))
                {
                    reason = "commentsExtensible requires unique durable IDs and boolean placeholder state";
                    return false;
                }
            }
            if (extensibleByDurableId.Count != comments.Count ||
                durableByParagraphId.Values.Any(id => !extensibleByDurableId.ContainsKey(id)))
            {
                reason = "commentsExtensible must map every durable comment exactly once";
                return false;
            }
        }
        foreach (var (paragraphId, entry) in extendedByParagraphId)
        {
            if (entry.Parent.Length == 0 || !durableByParagraphId.TryGetValue(paragraphId, out var durableId) ||
                !extensibleByDurableId.TryGetValue(durableId, out var extensible) || !extensible.Intelligent) continue;
            reason = $"reply comment paragraph {paragraphId} cannot be an intelligent placeholder";
            return false;
        }

        var peopleByAuthor = new Dictionary<string, DocumentCommentPerson>(StringComparer.Ordinal);
        if (peoplePart is not null)
        {
            var root = peoplePart.People;
            if (root is null || root.ChildElements.Any(child => child is not W15.Person))
            {
                reason = "people has no supported root or contains an unsupported child";
                return false;
            }
            foreach (var element in root.Elements<W15.Person>())
            {
                if (element.ChildElements.Any(child => child is not W15.PresenceInfo) || element.Elements<W15.PresenceInfo>().Count() != 1)
                {
                    reason = "each bounded comment person requires exactly one presenceInfo child";
                    return false;
                }
                var author = Attribute(element, "author", W15Ns) ?? string.Empty;
                var presence = element.Elements<W15.PresenceInfo>().Single();
                var providerId = Attribute(presence, "providerId", W15Ns) ?? string.Empty;
                var userId = Attribute(presence, "userId", W15Ns) ?? string.Empty;
                if (author.Length == 0 || providerId.Length is 0 or > 100 || userId.Length is 0 or > 300 ||
                    !peopleByAuthor.TryAdd(author, new DocumentCommentPerson { ProviderId = providerId, UserId = userId }))
                {
                    reason = "people requires unique non-empty authors with provider/user identity";
                    return false;
                }
            }
            var usedAuthors = comments.Select(comment => Attribute(comment, "author", "http://schemas.openxmlformats.org/wordprocessingml/2006/main") ?? string.Empty)
                .ToHashSet(StringComparer.Ordinal);
            if (peopleByAuthor.Keys.Any(author => !usedAuthors.Contains(author)))
            {
                reason = "people contains an author outside the bounded comment graph";
                return false;
            }
        }

        var byNativeId = new Dictionary<string, DocxExtendedCommentInfo>(StringComparer.Ordinal);
        foreach (var (paragraphId, source) in commentByParagraphId)
        {
            var extended = extendedByParagraphId[paragraphId];
            var durableId = durableByParagraphId.GetValueOrDefault(paragraphId, string.Empty);
            var extensible = durableId.Length > 0 && extensibleByDurableId.TryGetValue(durableId, out var value)
                ? value
                : (string.Empty, false);
            var author = Attribute(source.Element, "author", "http://schemas.openxmlformats.org/wordprocessingml/2006/main") ?? string.Empty;
            peopleByAuthor.TryGetValue(author, out var person);
            byNativeId.Add(source.NativeId, new DocxExtendedCommentInfo(
                paragraphId,
                extended.Parent,
                extended.Done,
                durableId,
                extensible.Item1,
                person?.Clone(),
                extensible.Item2));
        }

        graph = new DocxExtendedCommentGraph(
            true,
            byNativeId,
            commentsExPart,
            commentsIdsPart,
            commentsExtensiblePart,
            peoplePart,
            HashElement(commentsEx),
            commentsIdsPart?.CommentsIds is { } ids ? HashElement(ids) : string.Empty,
            commentsExtensiblePart?.CommentsExtensible is { } extensibleRoot ? HashElement(extensibleRoot) : string.Empty,
            peoplePart?.People is { } people ? HashElement(people) : string.Empty);
        return true;
    }

    internal static void Author(
        MainDocumentPart owner,
        IReadOnlyList<(DocumentComment Artifact, string NativeId, W.Comment Element)> comments)
    {
        var modern = comments.Any(item => IsModern(item.Artifact));
        if (!modern) return;

        var paragraphIds = PlanIds(
            comments.Select(item => (item.Artifact.Id, item.Artifact.ParagraphId)),
            "paragraph",
            boundedParagraph: true);
        var byArtifactId = comments.ToDictionary(item => item.Artifact.Id, StringComparer.Ordinal);
        var commentsExPart = owner.AddNewPart<WordprocessingCommentsExPart>();
        var commentsEx = new W15.CommentsEx();
        foreach (var item in comments)
        {
            var paragraphId = paragraphIds[item.Artifact.Id];
            SetAttribute(item.Element.Elements<W.Paragraph>().Single(), "w14", "paraId", W14, paragraphId);
            var element = new W15.CommentEx();
            SetAttribute(element, "w15", "paraId", W15Ns, paragraphId);
            if (item.Artifact.ParentCommentId.Length > 0)
            {
                var parent = byArtifactId[item.Artifact.ParentCommentId];
                SetAttribute(element, "w15", "paraIdParent", W15Ns, paragraphIds[parent.Artifact.Id]);
            }
            if (item.Artifact.HasResolved)
                SetAttribute(element, "w15", "done", W15Ns, item.Artifact.Resolved ? "1" : "0");
            commentsEx.Append(element);
        }
        commentsExPart.CommentsEx = commentsEx;
        commentsEx.Save(commentsExPart);

        var needsDurable = comments.Any(item => item.Artifact.DurableId.Length > 0 || item.Artifact.HasDateUtc || item.Artifact.HasIntelligentPlaceholder);
        Dictionary<string, string>? durableIds = null;
        if (needsDurable)
        {
            durableIds = PlanIds(comments.Select(item => (item.Artifact.Id, item.Artifact.DurableId)), "durable", boundedDurable: true);
            var idsPart = owner.AddNewPart<WordprocessingCommentsIdsPart>();
            var idsRoot = new W16Cid.CommentsIds();
            foreach (var item in comments)
            {
                var element = new W16Cid.CommentId();
                SetAttribute(element, "w16cid", "paraId", W16CidNs, paragraphIds[item.Artifact.Id]);
                SetAttribute(element, "w16cid", "durableId", W16CidNs, durableIds[item.Artifact.Id]);
                idsRoot.Append(element);
            }
            idsPart.CommentsIds = idsRoot;
            idsRoot.Save(idsPart);
        }

        if (comments.Any(item => item.Artifact.HasDateUtc || item.Artifact.HasIntelligentPlaceholder))
        {
            var extensiblePart = owner.AddNewPart<WordCommentsExtensiblePart>();
            var root = new W16Cex.CommentsExtensible();
            foreach (var item in comments)
            {
                var element = new W16Cex.CommentExtensible();
                SetAttribute(element, "w16cex", "durableId", W16CexNs, durableIds![item.Artifact.Id]);
                if (item.Artifact.HasDateUtc) SetAttribute(element, "w16cex", "dateUtc", W16CexNs, item.Artifact.DateUtc);
                if (item.Artifact.HasIntelligentPlaceholder)
                    SetAttribute(element, "w16cex", "intelligentPlaceholder", W16CexNs, item.Artifact.IntelligentPlaceholder ? "1" : "0");
                root.Append(element);
            }
            extensiblePart.CommentsExtensible = root;
            root.Save(extensiblePart);
        }

        var people = comments.Where(item => item.Artifact.Person is not null)
            .GroupBy(item => item.Artifact.Author, StringComparer.Ordinal)
            .Select(group => (Author: group.Key, Person: group.First().Artifact.Person))
            .OrderBy(item => item.Author, StringComparer.Ordinal)
            .ToArray();
        if (people.Length > 0)
        {
            var peoplePart = owner.AddNewPart<WordprocessingPeoplePart>();
            var root = new W15.People();
            foreach (var item in people)
            {
                var person = new W15.Person();
                SetAttribute(person, "w15", "author", W15Ns, item.Author);
                var presence = new W15.PresenceInfo();
                SetAttribute(presence, "w15", "providerId", W15Ns, item.Person.ProviderId);
                SetAttribute(presence, "w15", "userId", W15Ns, item.Person.UserId);
                person.Append(presence);
                root.Append(person);
            }
            peoplePart.People = root;
            root.Save(peoplePart);
        }
    }

    internal static bool ApplyResolved(
        DocxExtendedCommentGraph graph,
        string nativeCommentId,
        DocumentComment requested)
    {
        if (!graph.ByNativeCommentId.TryGetValue(nativeCommentId, out var original)) return false;
        if (!graph.IsModern)
        {
            if (requested.HasResolved)
                throw new CodecException("unsupported_document_comment_edit", "A legacy classic comment cannot gain modern resolved state in the fixed-topology edit slice.", "word/comments.xml");
            return false;
        }
        if (!requested.HasResolved)
            throw new CodecException("document_comment_source_binding_mismatch", "A modeled modern comment lost its resolved-state presence.", "word/commentsExtended.xml");
        if (requested.Resolved == original.Resolved) return false;
        var element = graph.CommentsExPart!.CommentsEx!.Elements<W15.CommentEx>()
            .Single(item => NormalizeParagraphId(Attribute(item, "paraId", W15Ns)) == original.ParagraphId);
        SetAttribute(element, "w15", "done", W15Ns, requested.Resolved ? "1" : "0");
        return true;
    }

    internal static void SaveResolved(DocxExtendedCommentGraph graph, DocxPartContext context)
    {
        var part = graph.CommentsExPart ?? throw new CodecException(
            "document_comment_source_binding_mismatch",
            "Modern comment state has no commentsExtended part.",
            "word/commentsExtended.xml");
        part.CommentsEx!.Save(part);
        context.MarkCommentSupportPartMutated(part);
    }

    private static bool IsModern(DocumentComment comment) =>
        comment.ParentCommentId.Length > 0 || comment.HasResolved || comment.ParagraphId.Length > 0 ||
        comment.DurableId.Length > 0 || comment.HasDateUtc || comment.Person is not null ||
        comment.HasIntelligentPlaceholder;

    private static Dictionary<string, string> PlanIds(
        IEnumerable<(string Id, string Requested)> source,
        string purpose,
        bool boundedDurable = false,
        bool boundedParagraph = false)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        var used = new HashSet<string>(StringComparer.Ordinal);
        foreach (var (id, requested) in source)
        {
            var normalized = boundedDurable
                ? NormalizeDurableId(requested)
                : boundedParagraph
                    ? NormalizeParagraphId(requested)
                    : NormalizeHex(requested);
            if (requested.Length > 0 && normalized is null)
                throw new CodecException(
                    "invalid_document_comment",
                    boundedDurable
                        ? $"Document comment {id} {purpose} ID must be between 00000001 and 7FFFFFFE."
                        : boundedParagraph
                            ? $"Document comment {id} {purpose} ID must be between 00000001 and 7FFFFFFF."
                        : $"Document comment {id} {purpose} ID must contain exactly eight hexadecimal digits.");
            var candidate = normalized;
            for (var salt = 0; candidate is null || used.Contains(candidate); salt++)
                candidate = DeterministicHex($"{purpose}:{id}:{salt}", boundedDurable, boundedParagraph);
            used.Add(candidate);
            result.Add(id, candidate);
        }
        return result;
    }

    private static string DeterministicHex(string value, bool boundedDurable, bool boundedParagraph)
    {
        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(value));
        if (!boundedDurable && !boundedParagraph) return Convert.ToHexString(hash.AsSpan(0, 4));
        var raw = BinaryPrimitives.ReadUInt32BigEndian(hash.AsSpan(0, 4));
        var number = boundedDurable ? raw % 0x7FFFFFFE + 1 : raw % 0x7FFFFFFF + 1;
        return number.ToString("X8", CultureInfo.InvariantCulture);
    }

    private static string? NormalizeHex(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length != 8 || value.Any(character => !Uri.IsHexDigit(character))) return null;
        return value.ToUpperInvariant();
    }

    private static string? NormalizeDurableId(string? value)
    {
        var normalized = NormalizeHex(value);
        if (normalized is null ||
            !uint.TryParse(normalized, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var number) ||
            number == 0 || number >= 0x7FFFFFFF) return null;
        return normalized;
    }

    private static string? NormalizeParagraphId(string? value)
    {
        var normalized = NormalizeHex(value);
        if (normalized is null ||
            !uint.TryParse(normalized, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var number) ||
            number == 0 || number >= 0x80000000) return null;
        return normalized;
    }

    private static string NormalizeOptionalParagraphId(string? value, out bool valid)
    {
        if (string.IsNullOrEmpty(value)) { valid = true; return string.Empty; }
        var normalized = NormalizeParagraphId(value);
        valid = normalized is not null;
        return normalized ?? string.Empty;
    }

    private static bool TryBoolean(string? value, out bool result)
    {
        if (string.IsNullOrEmpty(value) || value == "0" || value.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            result = false;
            return true;
        }
        if (value == "1" || value.Equals("true", StringComparison.OrdinalIgnoreCase))
        {
            result = true;
            return true;
        }
        result = false;
        return false;
    }

    private static string? Attribute(OpenXmlElement? element, string localName, string namespaceUri)
    {
        if (element is null) return null;
        var attribute = element.GetAttributes().FirstOrDefault(item => item.LocalName == localName && item.NamespaceUri == namespaceUri);
        return string.IsNullOrEmpty(attribute.LocalName) ? null : attribute.Value;
    }

    private static void SetAttribute(OpenXmlElement element, string prefix, string localName, string namespaceUri, string value) =>
        element.SetAttribute(new OpenXmlAttribute(prefix, localName, namespaceUri, value));

    private static string HashElement(OpenXmlElement element) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(element.OuterXml))).ToLowerInvariant();
}
