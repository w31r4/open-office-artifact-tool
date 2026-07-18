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

// Legacy PresentationML comments are intentionally kept separate from the
// richer JS thread model. The native format owns one author, one text value,
// and one slide coordinate per p:cm; it has no replies, resolved state,
// reactions, or element/text anchors. Modern comment parts stay opaque.
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

        var authorsPart = presentationPart.AddNewPart<CommentAuthorsPart>("rIdCommentAuthors");
        var authorList = new P.CommentAuthorList();
        foreach (var authorName in authorOrder)
        {
            var author = authorsByName[authorName];
            authorList.Append(new P.CommentAuthor
            {
                Id = author.NativeId,
                Name = author.Name,
                Initials = Initials(author.Name),
                LastIndex = author.LastIndex,
                ColorIndex = author.NativeId % 10,
            });
        }
        authorsPart.CommentAuthorList = authorList;

        // PresentationML comment indexes are one-based per author. Keep
        // lastIdx equal to the highest emitted index so package-level
        // semantic validation and host readers agree on the same invariant.
        var nextIndexByAuthor = authorsByName.Values.ToDictionary(author => author.Name, _ => 1U, StringComparer.Ordinal);
        for (var slideIndex = 0; slideIndex < slides.Count; slideIndex++)
        {
            var comments = slides[slideIndex].LegacyComments;
            if (comments.Count == 0) continue;
            var commentsPart = slideParts[slideIndex].AddNewPart<SlideCommentsPart>($"rIdComments{slideIndex + 1}");
            var commentList = new P.CommentList();
            foreach (var source in comments)
            {
                var author = authorsByName[source.Author];
                var nativeIndex = nextIndexByAuthor[source.Author]++;
                commentList.Append(new P.Comment(
                    new P.Position { X = source.PositionXEmu, Y = source.PositionYEmu },
                    new P.Text(source.Text))
                {
                    AuthorId = author.NativeId,
                    DateTime = ParseTimestamp(source.CreatedAt, $"Presentation slide {slideIndex + 1} legacy comment {source.Id}"),
                    Index = nativeIndex,
                });
            }
            commentsPart.CommentList = commentList;
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
            if (!Equivalent(profile.Comments, target))
                throw new CodecException(
                    "unsupported_presentation_comment_edit",
                    $"Presentation slide {slideIndex + 1} legacy comments are imported read-only; edit the source package with a specialized workflow or retain them unchanged.",
                    PartPath(slideParts[slideIndex]));
        }
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
    private static CodecException Invalid(string message) => new("invalid_presentation_legacy_comment", message);

    private sealed record RequestedComment(int SlideIndex, PresentationLegacyComment Comment);
    private sealed class AuthorBuild(uint nativeId, string name)
    {
        internal uint NativeId { get; } = nativeId;
        internal string Name { get; } = name;
        internal uint LastIndex { get; set; }
    }

}
