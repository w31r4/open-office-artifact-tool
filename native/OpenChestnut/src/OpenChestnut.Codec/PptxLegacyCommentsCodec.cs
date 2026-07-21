using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenChestnut.Codec;

internal sealed record PptxLegacyCommentProfile(bool Supported, IReadOnlyList<PresentationLegacyComment> Comments)
{
    internal static PptxLegacyCommentProfile Empty { get; } = new(true, []);
    internal static PptxLegacyCommentProfile Unsupported { get; } = new(false, []);
}

internal sealed record PptxLegacyCommentsChange(
    IReadOnlyList<string> ChangedPartPaths,
    IReadOnlyList<string> AddedPartPaths,
    IReadOnlyList<string> AddedRelationshipKeys,
    IReadOnlyDictionary<string, string> ReplacedPartHashes);

// Legacy PresentationML comments are intentionally kept separate from the
// richer JS thread model. The native format owns one author, one text value,
// and one slide coordinate per p:cm; it has no replies, resolved state,
// reactions, or element/text anchors. Office 2021 threads are handled by the
// separate bounded PptxModernCommentsCodec profile.
internal static class PptxLegacyCommentsCodec
{
    private const int MaxCommentsPerPresentation = 4_096;
    private const int MaxAuthorsPerPresentation = 256;
    private const int MaxAuthorLength = 255;
    private const int MaxCommentTextLength = 32_767;

    internal static IReadOnlyList<PresentationLegacyComment> Read(
        PresentationPart presentationPart,
        SlidePart slidePart,
        int slideIndex,
        IList<Diagnostic> diagnostics)
    {
        var profile = Profile(presentationPart, slidePart, slideIndex);
        if (profile.Supported) return profile.Comments;
        diagnostics.Add(CodecProtocol.Warning(
            "unsupported_presentation_legacy_comments_preserved",
            $"Presentation slide {slideIndex + 1} has a legacy comment graph outside the bounded single-text profile; it remains opaque and source-bound.",
            PartPath(slidePart)));
        return [];
    }

    // The first source-bound authoring slice deliberately starts from a deck
    // with no comment family at all. That gives the transaction exclusive
    // ownership of the new CommentAuthorsPart and every new SlideCommentsPart,
    // so an Agent cannot accidentally splice a simplified graph into an
    // existing legacy or Office 2021 review topology.
    internal static bool CanAddSourceBound(PresentationPart presentationPart, SlidePart slidePart)
    {
        if (presentationPart.Presentation?.SlideIdList is null ||
            !presentationPart.SlideParts.Contains(slidePart) ||
            presentationPart.Parts.Any(pair => pair.OpenXmlPart is CommentAuthorsPart or PowerPointAuthorsPart))
            return false;

        return presentationPart.SlideParts.All(candidate =>
            candidate.Parts.All(pair => pair.OpenXmlPart is not SlideCommentsPart and not PowerPointCommentPart));
    }

    internal static bool CommentPartPresent(SlidePart slidePart) =>
        slidePart.Parts.Any(pair => pair.OpenXmlPart is SlideCommentsPart or PowerPointCommentPart);

    internal static string CommentFamily(PresentationPart presentationPart)
    {
        var legacy = presentationPart.Parts.Any(pair => pair.OpenXmlPart is CommentAuthorsPart) ||
                     presentationPart.SlideParts.Any(slide => slide.Parts.Any(pair => pair.OpenXmlPart is SlideCommentsPart));
        var modern = presentationPart.Parts.Any(pair => pair.OpenXmlPart is PowerPointAuthorsPart) ||
                     presentationPart.SlideParts.Any(slide => slide.Parts.Any(pair => pair.OpenXmlPart is PowerPointCommentPart));
        return (legacy, modern) switch
        {
            (true, true) => "mixed",
            (true, false) => "legacy",
            (false, true) => "modern",
            _ => "legacy",
        };
    }

    internal static void BuildSourceFree(
        PresentationPart presentationPart,
        IReadOnlyList<SlidePart> slideParts,
        IReadOnlyList<PresentationSlide> slides)
    {
        var requested = slides
            .SelectMany((slide, slideIndex) => slide.LegacyComments.Select(comment => new RequestedComment(slideIndex, comment)))
            .ToArray();
        if (requested.Length == 0) return;
        if (requested.Length > MaxCommentsPerPresentation)
            throw Invalid("Presentation legacy-comment count exceeds the 4096-comment budget.");

        var (authorOrder, authorsByName) = CollectAuthors(requested);

        var authorsPart = presentationPart.AddNewPart<CommentAuthorsPart>("rIdCommentAuthors");
        authorsPart.CommentAuthorList = AuthorList(authorOrder, authorsByName);

        // PresentationML comment indexes are one-based per author. Keep
        // lastIdx equal to the highest emitted index so package-level
        // semantic validation and host readers agree on the same invariant.
        var nextIndexByAuthor = authorsByName.Values.ToDictionary(author => author.Name, _ => 1U, StringComparer.Ordinal);
        for (var slideIndex = 0; slideIndex < slides.Count; slideIndex++)
        {
            var comments = slides[slideIndex].LegacyComments;
            if (comments.Count == 0) continue;
            var commentsPart = slideParts[slideIndex].AddNewPart<SlideCommentsPart>($"rIdComments{slideIndex + 1}");
            commentsPart.CommentList = CommentList(comments, slideIndex, authorsByName, nextIndexByAuthor);
        }
    }

    internal static void AssertSourceUnchanged(
        PresentationPart presentationPart,
        IReadOnlyList<SlidePart> slideParts,
        IReadOnlyList<PresentationSlide> requested)
    {
        if (slideParts.Count != requested.Count)
            throw new CodecException("presentation_comment_topology_changed", "Presentation slide topology changed before legacy comments could be verified.");
        for (var slideIndex = 0; slideIndex < slideParts.Count; slideIndex++)
        {
            var targetSlide = requested[slideIndex];
            // Reordering keeps the original SlidePart and its package-local
            // comment locator. Use the source index for identity comparison,
            // not the presentation's current display order.
            var sourceIndex = targetSlide.Source?.SlideIndex is { } boundIndex
                ? checked((int)boundIndex)
                : slideIndex;
            var profile = Profile(presentationPart, slideParts[slideIndex], sourceIndex);
            var target = targetSlide.LegacyComments;
            if (!profile.Supported)
            {
                if (target.Count > 0)
                    throw new CodecException(
                        "unsupported_presentation_comment_edit",
                        $"Presentation slide {slideIndex + 1} has an unmodeled legacy comment graph and cannot add or replace comments through the bounded profile.",
                        PartPath(slideParts[slideIndex]));
                continue;
            }
            if (Equivalent(profile.Comments, target)) continue;
            if (profile.Comments.Count == 0 && target.Count > 0 && CanAddSourceBound(presentationPart, slideParts[slideIndex]))
            {
                foreach (var comment in target)
                {
                    Validate(comment, sourceIndex);
                    if (comment.NativeAuthorId != 0 || comment.NativeIndex != 0)
                        throw new CodecException(
                            "presentation_comment_topology_changed",
                            $"Presentation slide {slideIndex + 1} new legacy comments cannot claim package-local author or comment indexes.",
                            PartPath(slideParts[slideIndex]));
                }
                continue;
            }
            throw new CodecException(
                "unsupported_presentation_comment_edit",
                $"Presentation slide {slideIndex + 1} legacy comments are imported read-only; edit the source package with a specialized workflow or retain them unchanged.",
                PartPath(slideParts[slideIndex]));
        }
    }

    internal static PptxLegacyCommentsChange? ApplySourceBoundAdditions(
        PresentationPart presentationPart,
        IReadOnlyList<SlidePart> slideParts,
        IReadOnlyList<PresentationSlide> slides)
    {
        if (slideParts.Count != slides.Count)
            throw new CodecException("presentation_comment_topology_changed", "Presentation slide topology changed before legacy comments could be added.");

        var requested = new List<RequestedComment>();
        for (var slideIndex = 0; slideIndex < slides.Count; slideIndex++)
        {
            var sourceIndex = slides[slideIndex].Source?.SlideIndex is { } boundIndex
                ? checked((int)boundIndex)
                : slideIndex;
            var profile = Profile(presentationPart, slideParts[slideIndex], sourceIndex);
            if (!profile.Supported || profile.Comments.Count > 0 || slides[slideIndex].LegacyComments.Count == 0) continue;
            if (!CanAddSourceBound(presentationPart, slideParts[slideIndex]))
                throw new CodecException(
                    "unsupported_presentation_comment_edit",
                    $"Imported presentation slide {slideIndex + 1} cannot add a legacy comments part because the source presentation already owns an incompatible comment graph.",
                    PartPath(slideParts[slideIndex]));
            requested.AddRange(slides[slideIndex].LegacyComments.Select(comment => new RequestedComment(slideIndex, comment)));
        }
        if (requested.Count == 0) return null;
        if (requested.Count > MaxCommentsPerPresentation)
            throw Invalid("Presentation legacy-comment count exceeds the 4096-comment budget.");

        // Re-prove every candidate before the first package mutation. Once the
        // shared author catalog exists, CanAddSourceBound must become false.
        foreach (var slideIndex in requested.Select(item => item.SlideIndex).Distinct())
            if (!CanAddSourceBound(presentationPart, slideParts[slideIndex]))
                throw new CodecException(
                    "unsupported_presentation_comment_edit",
                    $"Imported presentation slide {slideIndex + 1} cannot add a legacy comments part.",
                    PartPath(slideParts[slideIndex]));

        var (authorOrder, authorsByName) = CollectAuthors(requested);
        var changedPartPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            RelationshipPartPath(presentationPart),
            "[Content_Types].xml",
        };
        var addedPartPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var addedRelationshipKeys = new HashSet<string>(StringComparer.Ordinal);

        var authorsRelationshipId = NextRelationshipId(presentationPart, "rIdCommentAuthors");
        var authorsPart = presentationPart.AddNewPart<CommentAuthorsPart>(authorsRelationshipId);
        authorsPart.CommentAuthorList = AuthorList(authorOrder, authorsByName);
        authorsPart.CommentAuthorList.Save();
        changedPartPaths.Add(PartPath(authorsPart));
        addedPartPaths.Add(PartPath(authorsPart));
        addedRelationshipKeys.Add(RelationshipKey(presentationPart, authorsRelationshipId));

        var nextIndexByAuthor = authorsByName.Values.ToDictionary(author => author.Name, _ => 1U, StringComparer.Ordinal);
        foreach (var slideIndex in requested.Select(item => item.SlideIndex).Distinct().Order())
        {
            var slidePart = slideParts[slideIndex];
            var relationshipId = NextRelationshipId(slidePart, "rIdComments");
            var commentsPart = slidePart.AddNewPart<SlideCommentsPart>(relationshipId);
            commentsPart.CommentList = CommentList(slides[slideIndex].LegacyComments, slideIndex, authorsByName, nextIndexByAuthor);
            commentsPart.CommentList.Save();
            changedPartPaths.Add(RelationshipPartPath(slidePart));
            changedPartPaths.Add(PartPath(commentsPart));
            addedPartPaths.Add(PartPath(commentsPart));
            addedRelationshipKeys.Add(RelationshipKey(slidePart, relationshipId));
        }

        return new PptxLegacyCommentsChange(
            changedPartPaths.ToArray(),
            addedPartPaths.ToArray(),
            addedRelationshipKeys.ToArray(),
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase));
    }

    internal static void ValidateSourceBoundOutput(
        PresentationPart sourcePresentationPart,
        PresentationPart outputPresentationPart,
        SlidePart sourceSlidePart,
        SlidePart outputSlidePart,
        PresentationSlide requested,
        int slideIndex)
    {
        var sourceIndex = requested.Source?.SlideIndex is { } boundIndex
            ? checked((int)boundIndex)
            : slideIndex;
        var source = Profile(sourcePresentationPart, sourceSlidePart, sourceIndex);
        var output = Profile(outputPresentationPart, outputSlidePart, sourceIndex);
        if (!source.Supported)
        {
            if (requested.LegacyComments.Count > 0)
                throw Postwrite(slideIndex, "an unsupported source legacy-comment graph was projected as modeled comments", PartPath(outputSlidePart));
            return;
        }
        if (!output.Supported)
            throw Postwrite(slideIndex, "the output legacy-comment graph is outside the bounded profile", PartPath(outputSlidePart));

        if (source.Comments.Count > 0)
        {
            if (!Equivalent(source.Comments, requested.LegacyComments) || !Equivalent(output.Comments, requested.LegacyComments))
                throw Postwrite(slideIndex, "an existing legacy comment changed", PartPath(outputSlidePart));
            return;
        }

        if (requested.LegacyComments.Count == 0)
        {
            if (output.Comments.Count != 0)
                throw Postwrite(slideIndex, "an unchanged slide unexpectedly gained legacy comments", PartPath(outputSlidePart));
            return;
        }

        if (!CanAddSourceBound(sourcePresentationPart, sourceSlidePart) ||
            !EquivalentAdded(output.Comments, requested.LegacyComments))
            throw Postwrite(slideIndex, "the added legacy comments do not match the requested artifact", PartPath(outputSlidePart));
        ValidateAddedGraph(outputPresentationPart, outputSlidePart, slideIndex);
    }

    internal static void Validate(PresentationSlide slide, int slideIndex)
    {
        if (slide.LegacyComments.Count > MaxCommentsPerPresentation)
            throw Invalid($"Presentation slide {slideIndex + 1} exceeds the 4096-comment budget.");
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var comment in slide.LegacyComments)
        {
            Validate(comment, slideIndex);
            if (!ids.Add(comment.Id))
                throw Invalid($"Presentation slide {slideIndex + 1} has duplicate legacy comment ID {comment.Id}.");
        }
    }

    // The clone preflight uses this same parser instead of maintaining an
    // independent relationship/profile inventory. Its result is deliberately
    // internal: legacy author/index values are package-local evidence, not a
    // public cross-file identity contract.
    internal static PptxLegacyCommentProfile Profile(PresentationPart presentationPart, SlidePart slidePart, int slideIndex)
    {
        var commentsPart = slidePart.SlideCommentsPart;
        if (commentsPart is null) return PptxLegacyCommentProfile.Empty;
        var authorList = presentationPart.CommentAuthorsPart?.CommentAuthorList;
        var commentList = commentsPart.CommentList;
        if (authorList is null || commentList is null ||
            authorList.ChildElements.Any(item => item is not P.CommentAuthor) ||
            commentList.ChildElements.Any(item => item is not P.Comment)) return PptxLegacyCommentProfile.Unsupported;

        var authors = new Dictionary<uint, string>();
        foreach (var author in authorList.Elements<P.CommentAuthor>())
        {
            if (author.ChildElements.Count != 0 || author.Id?.Value is not uint nativeId ||
                string.IsNullOrWhiteSpace(author.Name?.Value) || author.Name.Value.Length > MaxAuthorLength ||
                !authors.TryAdd(nativeId, author.Name.Value)) return PptxLegacyCommentProfile.Unsupported;
        }
        if (commentList.Elements<P.Comment>().Count() > MaxCommentsPerPresentation) return PptxLegacyCommentProfile.Unsupported;

        var comments = new List<PresentationLegacyComment>();
        var nativeIds = new HashSet<(uint AuthorId, uint Index)>();
        foreach (var comment in commentList.Elements<P.Comment>())
        {
            if (comment.ChildElements.Count != 2 || comment.Position is null || comment.Text is null ||
                comment.Position.X?.Value is not long x || comment.Position.Y?.Value is not long y ||
                comment.AuthorId?.Value is not uint authorId || comment.Index?.Value is not uint nativeIndex || nativeIndex == 0 ||
                comment.DateTime?.Value is not DateTime created ||
                comment.Text.Text is not { } text || text.Length > MaxCommentTextLength ||
                !authors.TryGetValue(authorId, out var author) || !nativeIds.Add((authorId, nativeIndex)))
                return PptxLegacyCommentProfile.Unsupported;
            comments.Add(new PresentationLegacyComment
            {
                Id = $"presentation/slide/{slideIndex + 1}/legacy-comment/{comments.Count + 1}",
                Author = author,
                Text = text,
                CreatedAt = created.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture),
                PositionXEmu = x,
                PositionYEmu = y,
                NativeAuthorId = authorId,
                NativeIndex = nativeIndex,
            });
        }
        return new PptxLegacyCommentProfile(true, comments);
    }

    internal static bool Equivalent(IReadOnlyList<PresentationLegacyComment> actual, IList<PresentationLegacyComment> requested) =>
        actual.Count == requested.Count && actual.Zip(requested).All(pair =>
            pair.First.Id == pair.Second.Id &&
            pair.First.Author == pair.Second.Author &&
            pair.First.Text == pair.Second.Text &&
            pair.First.CreatedAt == pair.Second.CreatedAt &&
            pair.First.PositionXEmu == pair.Second.PositionXEmu &&
            pair.First.PositionYEmu == pair.Second.PositionYEmu &&
            pair.First.NativeAuthorId == pair.Second.NativeAuthorId &&
            pair.First.NativeIndex == pair.Second.NativeIndex);

    private static bool EquivalentAdded(IReadOnlyList<PresentationLegacyComment> actual, IList<PresentationLegacyComment> requested) =>
        actual.Count == requested.Count && actual.Zip(requested).All(pair =>
            pair.First.Author == pair.Second.Author &&
            pair.First.Text == pair.Second.Text &&
            pair.First.CreatedAt == ParseTimestamp(pair.Second.CreatedAt, $"Added legacy comment {pair.Second.Id}").ToUniversalTime().ToString("O", CultureInfo.InvariantCulture) &&
            pair.First.PositionXEmu == pair.Second.PositionXEmu &&
            pair.First.PositionYEmu == pair.Second.PositionYEmu &&
            pair.First.NativeIndex > 0);

    private static (IReadOnlyList<string> AuthorOrder, Dictionary<string, AuthorBuild> AuthorsByName) CollectAuthors(
        IReadOnlyCollection<RequestedComment> requested)
    {
        var authorOrder = new List<string>();
        var authorsByName = new Dictionary<string, AuthorBuild>(StringComparer.Ordinal);
        foreach (var item in requested)
        {
            Validate(item.Comment, item.SlideIndex);
            if (!authorsByName.TryGetValue(item.Comment.Author, out var author))
            {
                if (authorOrder.Count >= MaxAuthorsPerPresentation)
                    throw Invalid("Presentation legacy-comment profile exceeds the 256-author budget.");
                author = new AuthorBuild(checked((uint)authorOrder.Count), item.Comment.Author);
                authorsByName.Add(item.Comment.Author, author);
                authorOrder.Add(item.Comment.Author);
            }
            author.LastIndex = checked(author.LastIndex + 1);
        }
        return (authorOrder, authorsByName);
    }

    private static P.CommentAuthorList AuthorList(
        IEnumerable<string> authorOrder,
        IReadOnlyDictionary<string, AuthorBuild> authorsByName)
    {
        var list = new P.CommentAuthorList();
        foreach (var authorName in authorOrder)
        {
            var author = authorsByName[authorName];
            list.Append(new P.CommentAuthor
            {
                Id = author.NativeId,
                Name = author.Name,
                Initials = Initials(author.Name),
                LastIndex = author.LastIndex,
                ColorIndex = author.NativeId % 10,
            });
        }
        return list;
    }

    private static P.CommentList CommentList(
        IEnumerable<PresentationLegacyComment> comments,
        int slideIndex,
        IReadOnlyDictionary<string, AuthorBuild> authorsByName,
        IDictionary<string, uint> nextIndexByAuthor)
    {
        var list = new P.CommentList();
        foreach (var source in comments)
        {
            var author = authorsByName[source.Author];
            var nativeIndex = nextIndexByAuthor[source.Author]++;
            list.Append(new P.Comment(
                new P.Position { X = source.PositionXEmu, Y = source.PositionYEmu },
                new P.Text(source.Text))
            {
                AuthorId = author.NativeId,
                DateTime = ParseTimestamp(source.CreatedAt, $"Presentation slide {slideIndex + 1} legacy comment {source.Id}"),
                Index = nativeIndex,
            });
        }
        return list;
    }

    private static void ValidateAddedGraph(PresentationPart presentationPart, SlidePart slidePart, int slideIndex)
    {
        var authorsParts = presentationPart.Parts.Where(pair => pair.OpenXmlPart is CommentAuthorsPart).ToArray();
        var modernAuthors = presentationPart.Parts.Any(pair => pair.OpenXmlPart is PowerPointAuthorsPart);
        var commentParts = slidePart.Parts.Where(pair => pair.OpenXmlPart is SlideCommentsPart).ToArray();
        var modernComments = presentationPart.SlideParts.Any(candidate => candidate.Parts.Any(pair => pair.OpenXmlPart is PowerPointCommentPart));
        if (authorsParts.Length != 1 || modernAuthors || commentParts.Length != 1 || modernComments)
            throw Postwrite(slideIndex, "the added legacy comment graph is missing, ambiguous, or mixed with Office 2021 comments", PartPath(slidePart));

        var authorsPart = (CommentAuthorsPart)authorsParts[0].OpenXmlPart;
        var commentsPart = (SlideCommentsPart)commentParts[0].OpenXmlPart;
        if (HasRelationships(authorsPart) || HasRelationships(commentsPart) ||
            authorsPart.CommentAuthorList is not { } authorList ||
            commentsPart.CommentList is not { })
            throw Postwrite(slideIndex, "the added legacy comment graph is connected or incomplete", PartPath(slidePart));

        var indexesByAuthor = new Dictionary<uint, List<uint>>();
        foreach (var candidate in presentationPart.SlideParts)
        {
            var part = candidate.SlideCommentsPart;
            if (part is null) continue;
            if (HasRelationships(part) || part.CommentList is not { } list)
                throw Postwrite(slideIndex, "an added SlideCommentsPart is connected or incomplete", PartPath(part));
            foreach (var comment in list.Elements<P.Comment>())
            {
                if (comment.AuthorId?.Value is not uint authorId || comment.Index?.Value is not uint index || index == 0)
                    throw Postwrite(slideIndex, "an added comment has no canonical author/index identity", PartPath(part));
                if (!indexesByAuthor.TryGetValue(authorId, out var indexes))
                {
                    indexes = [];
                    indexesByAuthor.Add(authorId, indexes);
                }
                indexes.Add(index);
            }
        }

        var nativeIds = new HashSet<uint>();
        foreach (var author in authorList.Elements<P.CommentAuthor>())
        {
            if (author.ChildElements.Count != 0 || author.Id?.Value is not uint id ||
                string.IsNullOrWhiteSpace(author.Name?.Value) || !nativeIds.Add(id) ||
                author.Initials?.Value != Initials(author.Name.Value) ||
                author.ColorIndex?.Value != id % 10 ||
                !indexesByAuthor.TryGetValue(id, out var indexes) ||
                !indexes.Order().SequenceEqual(Enumerable.Range(1, indexes.Count).Select(index => checked((uint)index))) ||
                author.LastIndex?.Value != checked((uint)indexes.Count))
                throw Postwrite(slideIndex, "the added legacy author catalog is not canonical", PartPath(authorsPart));
        }
        if (nativeIds.Count == 0 || indexesByAuthor.Keys.Any(id => !nativeIds.Contains(id)))
            throw Postwrite(slideIndex, "the added legacy comments do not resolve through the author catalog", PartPath(authorsPart));
    }

    private static void Validate(PresentationLegacyComment comment, int slideIndex)
    {
        if (string.IsNullOrWhiteSpace(comment.Id) || comment.Author.Length is 0 or > MaxAuthorLength ||
            comment.Text.Length > MaxCommentTextLength)
            throw Invalid($"Presentation slide {slideIndex + 1} has invalid legacy-comment identity, author, or text.");
        _ = ParseTimestamp(comment.CreatedAt, $"Presentation slide {slideIndex + 1} legacy comment {comment.Id}");
    }

    private static DateTime ParseTimestamp(string value, string context)
    {
        if (!DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var timestamp))
            throw Invalid($"{context} has an invalid ISO-8601 created_at value.");
        return timestamp.UtcDateTime;
    }

    private static string Initials(string author)
    {
        var initials = string.Concat(author
            .Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .Take(3)
            .Select(part => char.ToUpperInvariant(part[0])));
        return initials.Length == 0 ? "OA" : initials;
    }

    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static bool HasRelationships(OpenXmlPart part) =>
        part.Parts.Any() || part.ExternalRelationships.Any() || part.HyperlinkRelationships.Any() || part.DataPartReferenceRelationships.Any();
    private static string RelationshipKey(OpenXmlPart source, string relationshipId) => $"{PartPath(source)}\0{relationshipId}";
    private static string RelationshipPartPath(OpenXmlPart part)
    {
        var path = PartPath(part);
        var separator = path.LastIndexOf('/');
        var directory = separator < 0 ? string.Empty : path[..separator];
        var fileName = separator < 0 ? path : path[(separator + 1)..];
        return directory.Length == 0 ? $"_rels/{fileName}.rels" : $"{directory}/_rels/{fileName}.rels";
    }
    private static string NextRelationshipId(OpenXmlPartContainer owner, string stem)
    {
        var used = owner.Parts.Select(pair => pair.RelationshipId)
            .Concat(owner.ExternalRelationships.Select(relationship => relationship.Id))
            .Concat(owner.HyperlinkRelationships.Select(relationship => relationship.Id))
            .Concat(owner.DataPartReferenceRelationships.Select(relationship => relationship.Id))
            .ToHashSet(StringComparer.Ordinal);
        for (var index = 1; index <= 1_000_000; index++)
        {
            var candidate = stem + index;
            if (!used.Contains(candidate)) return candidate;
        }
        throw new CodecException("presentation_relationship_budget_exceeded", "PPTX relationship ID allocation exceeded its bounded search.");
    }
    private static CodecException Postwrite(int slideIndex, string message, string path) =>
        new("presentation_postwrite_comment_mismatch", $"PPTX slide {slideIndex + 1} {message}.", path);
    private static CodecException Invalid(string message) => new("invalid_presentation_legacy_comment", message);

    private sealed record RequestedComment(int SlideIndex, PresentationLegacyComment Comment);
    private sealed class AuthorBuild(uint nativeId, string name)
    {
        internal uint NativeId { get; } = nativeId;
        internal string Name { get; } = name;
        internal uint LastIndex { get; set; }
    }

}
