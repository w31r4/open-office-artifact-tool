using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using A16 = DocumentFormat.OpenXml.Office2016.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenChestnut.Codec;

internal sealed record PptxModernCommentsChange(string PartPath, string Sha256);

// Office 2021 modern comments are a different package graph from the legacy
// p:cm format. This codec deliberately owns a narrow, auditable profile:
// one presentation-wide author catalog, one closed comments part per slide,
// one root plus direct replies, plain text, and one top-level drawing or shape
// text-range anchor. Rich task fields, reactions, extensions, unknown anchors,
// nested moniker chains, and connected comment parts remain opaque/source-bound.
internal static class PptxModernCommentsCodec
{
    private const string CommentNamespace = "http://schemas.microsoft.com/office/powerpoint/2018/8/main";
    private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string DrawingCommandNamespace = "http://schemas.microsoft.com/office/drawing/2013/main/command";
    private const string PresentationCommandNamespace = "http://schemas.microsoft.com/office/powerpoint/2013/main/command";
    private const int MaxCommentsPerPresentation = 4_096;
    private const int MaxAuthorsPerPresentation = 256;
    private const int MaxTextLength = 32_767;
    private const int MaxMetadataLength = 1_024;

    private static readonly XNamespace P188 = CommentNamespace;
    private static readonly XNamespace Drawing = DrawingNamespace;
    private static readonly XNamespace Oac = DrawingCommandNamespace;
    private static readonly XNamespace Pc = PresentationCommandNamespace;
    private static readonly XNamespace Xml = XNamespace.Xml;
    private static readonly HashSet<string> Statuses = new(StringComparer.Ordinal) { "active", "resolved", "closed" };
    private static readonly HashSet<string> Monikers = new(StringComparer.Ordinal) { "spMk", "graphicFrameMk", "cxnSpMk", "picMk", "grpSpMk" };

    internal static IReadOnlyList<PresentationModernCommentThread> Read(
        PresentationPart presentationPart,
        P.SlideId slideId,
        SlidePart slidePart,
        IReadOnlyList<OpenXmlElement> elements,
        IReadOnlyDictionary<uint, string> elementIdsByNativeId,
        int slideIndex,
        IList<Diagnostic> diagnostics)
    {
        var commentParts = Parts<PowerPointCommentPart>(slidePart);
        if (commentParts.Length == 0) return [];
        if (commentParts.Length != 1 || slidePart.SlideCommentsPart is not null)
            return Unsupported(slidePart, slideIndex, diagnostics, "multiple or mixed legacy/modern comment parts");

        var authorParts = Parts<PowerPointAuthorsPart>(presentationPart);
        if (authorParts.Length != 1)
            return Unsupported(slidePart, slideIndex, diagnostics, "a missing or ambiguous modern author catalog");

        var commentPart = commentParts[0];
        var authorsPart = authorParts[0];
        if (HasRelationships(commentPart) || HasRelationships(authorsPart))
            return Unsupported(slidePart, slideIndex, diagnostics, "connected modern comment or author parts");

        try
        {
            var authorBytes = ReadBytes(authorsPart);
            var commentBytes = ReadBytes(commentPart);
            var authors = ParseAuthors(authorBytes);
            var targets = AnchorTargets(elements, elementIdsByNativeId);
            var nativeSlideId = slideId.Id?.Value ?? 0U;
            var threads = ParseComments(commentBytes, authors, targets, nativeSlideId, slideIndex);
            var commentHash = Hash(commentBytes);
            var authorsHash = Hash(authorBytes);
            var commentRelationshipId = slidePart.GetIdOfPart(commentPart);
            var authorsRelationshipId = presentationPart.GetIdOfPart(authorsPart);
            for (var rootIndex = 0; rootIndex < threads.Count; rootIndex++)
            {
                var thread = threads[rootIndex];
                thread.Source = new PresentationModernCommentSourceBinding
                {
                    PartPath = PartPath(commentPart),
                    RelationshipId = commentRelationshipId,
                    CommentXmlSha256 = commentHash,
                    AuthorsPartPath = PartPath(authorsPart),
                    AuthorsRelationshipId = authorsRelationshipId,
                    AuthorsXmlSha256 = authorsHash,
                    FixedTopologySha256 = FixedTopologyHash(thread),
                    RootIndex = checked((uint)rootIndex),
                    Editable = true,
                };
            }
            return threads;
        }
        catch (UnsupportedProfileException exception)
        {
            return Unsupported(slidePart, slideIndex, diagnostics, exception.Message);
        }
    }

    internal static void BuildSourceFree(
        PresentationPart presentationPart,
        IReadOnlyList<P.SlideId> slideIds,
        IReadOnlyList<SlidePart> slideParts,
        IReadOnlyList<PresentationSlide> slides)
    {
        var requested = slides.SelectMany(slide => slide.ModernComments).ToArray();
        if (requested.Length == 0) return;
        if (slides.Any(slide => slide.LegacyComments.Count > 0))
            throw Invalid("A presentation cannot mix legacy and modern comment wire families.");
        if (requested.Length > MaxCommentsPerPresentation)
            throw Invalid("Presentation modern-comment count exceeds the 4096-thread budget.");

        var authors = CollectAuthors(requested);
        var authorsPart = presentationPart.AddNewPart<PowerPointAuthorsPart>("rIdModernCommentAuthors");
        WriteDocument(authorsPart, AuthorsDocument(authors));

        for (var slideIndex = 0; slideIndex < slides.Count; slideIndex++)
        {
            var source = slides[slideIndex];
            if (source.ModernComments.Count == 0) continue;
            var nativeSlideId = slideIds[slideIndex].Id?.Value ?? 0U;
            foreach (var thread in source.ModernComments) Validate(thread, source, nativeSlideId, slideIndex, sourceBound: false);
            var commentsPart = slideParts[slideIndex].AddNewPart<PowerPointCommentPart>($"rIdModernComments{slideIndex + 1}");
            WriteDocument(commentsPart, CommentsDocument(source.ModernComments));
        }
    }

    internal static PptxModernCommentsChange? ApplySourceBound(
        PresentationPart presentationPart,
        P.SlideId slideId,
        SlidePart slidePart,
        IReadOnlyList<OpenXmlElement> elements,
        IReadOnlyDictionary<uint, string> elementIdsByNativeId,
        PresentationSlide requested,
        int slideIndex)
    {
        var commentParts = Parts<PowerPointCommentPart>(slidePart);
        if (commentParts.Length == 0)
        {
            if (requested.ModernComments.Count > 0)
                throw new CodecException("unsupported_presentation_comment_edit", $"Imported presentation slide {slideIndex + 1} cannot add a modern comment part.", PartPath(slidePart));
            return null;
        }
        if (commentParts.Length != 1 || slidePart.SlideCommentsPart is not null)
            return RejectModeledEdit(requested, slidePart, slideIndex, "mixed or ambiguous comment parts");
        var authorParts = Parts<PowerPointAuthorsPart>(presentationPart);
        if (authorParts.Length != 1)
            return RejectModeledEdit(requested, slidePart, slideIndex, "a missing or ambiguous author catalog");

        var commentPart = commentParts[0];
        var authorsPart = authorParts[0];
        if (HasRelationships(commentPart) || HasRelationships(authorsPart))
            return RejectModeledEdit(requested, slidePart, slideIndex, "connected comment or author parts");

        try
        {
            var authorBytes = ReadBytes(authorsPart);
            var commentBytes = ReadBytes(commentPart);
            var authors = ParseAuthors(authorBytes);
            var targets = AnchorTargets(elements, elementIdsByNativeId);
            var nativeSlideId = slideId.Id?.Value ?? 0U;
            var sourceThreads = ParseComments(commentBytes, authors, targets, nativeSlideId, slideIndex);
            if (sourceThreads.Count != requested.ModernComments.Count)
                throw TopologyChanged(slideIndex, commentPart);

            var commentHash = Hash(commentBytes);
            var authorsHash = Hash(authorBytes);
            var partPath = PartPath(commentPart);
            var authorsPath = PartPath(authorsPart);
            var commentRelationshipId = slidePart.GetIdOfPart(commentPart);
            var authorsRelationshipId = presentationPart.GetIdOfPart(authorsPart);
            var document = LoadDocument(commentBytes);
            var roots = document.Root!.Elements(P188 + "cm").ToArray();
            var changed = false;

            for (var rootIndex = 0; rootIndex < sourceThreads.Count; rootIndex++)
            {
                var original = sourceThreads[rootIndex];
                var target = requested.ModernComments[rootIndex];
                var binding = target.Source;
                if (binding is null || !binding.Editable || binding.RootIndex != rootIndex ||
                    !binding.PartPath.Equals(partPath, StringComparison.OrdinalIgnoreCase) ||
                    binding.RelationshipId != commentRelationshipId ||
                    !binding.CommentXmlSha256.Equals(commentHash, StringComparison.OrdinalIgnoreCase) ||
                    !binding.AuthorsPartPath.Equals(authorsPath, StringComparison.OrdinalIgnoreCase) ||
                    binding.AuthorsRelationshipId != authorsRelationshipId ||
                    !binding.AuthorsXmlSha256.Equals(authorsHash, StringComparison.OrdinalIgnoreCase) ||
                    !binding.FixedTopologySha256.Equals(FixedTopologyHash(original), StringComparison.OrdinalIgnoreCase) ||
                    !binding.FixedTopologySha256.Equals(FixedTopologyHash(target), StringComparison.OrdinalIgnoreCase))
                    throw TopologyChanged(slideIndex, commentPart);

                Validate(target, requested, nativeSlideId, slideIndex, sourceBound: true);
                var rootElement = roots[rootIndex];
                changed |= ApplyMutableComment(rootElement, original.Root, target.Root);
                var replyElements = rootElement.Element(P188 + "replyLst")?.Elements(P188 + "reply").ToArray() ?? [];
                if (replyElements.Length != original.Replies.Count || replyElements.Length != target.Replies.Count)
                    throw TopologyChanged(slideIndex, commentPart);
                for (var replyIndex = 0; replyIndex < replyElements.Length; replyIndex++)
                    changed |= ApplyMutableComment(replyElements[replyIndex], original.Replies[replyIndex], target.Replies[replyIndex]);
            }

            if (!changed) return null;
            var output = DocumentBytes(document);
            WriteBytes(commentPart, output);
            return new PptxModernCommentsChange(partPath, Hash(output));
        }
        catch (UnsupportedProfileException exception)
        {
            return RejectModeledEdit(requested, slidePart, slideIndex, exception.Message);
        }
    }

    internal static void Validate(PresentationSlide slide, int slideIndex, bool hasSourcePackage)
    {
        if (slide.LegacyComments.Count > 0 && slide.ModernComments.Count > 0)
            throw Invalid($"Presentation slide {slideIndex + 1} mixes legacy and modern comments.");
        if (slide.ModernComments.Count > MaxCommentsPerPresentation)
            throw Invalid($"Presentation slide {slideIndex + 1} exceeds the 4096-thread modern-comment budget.");
        var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var thread in slide.ModernComments)
        {
            Validate(thread, slide, thread.Anchor?.NativeSlideId ?? 0, slideIndex, hasSourcePackage);
            foreach (var comment in new[] { thread.Root }.Concat(thread.Replies))
                if (!ids.Add(comment.Id)) throw Invalid($"Presentation slide {slideIndex + 1} has duplicate modern-comment ID {comment.Id}.");
        }
    }

    private static IReadOnlyList<PresentationModernCommentThread> ParseComments(
        byte[] bytes,
        IReadOnlyDictionary<string, PresentationModernComment> authors,
        IReadOnlyDictionary<uint, AnchorTarget> targets,
        uint nativeSlideId,
        int slideIndex)
    {
        var document = LoadDocument(bytes);
        var root = document.Root;
        if (root?.Name != P188 + "cmLst" || SignificantAttributes(root).Any() || root.Elements().Any(element => element.Name != P188 + "cm"))
            throw Unsupported("the comment list root or children are outside the bounded profile");
        var roots = root.Elements(P188 + "cm").ToArray();
        if (roots.Length > MaxCommentsPerPresentation) throw Unsupported("the comment count exceeds the bounded profile budget");
        var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var output = new List<PresentationModernCommentThread>();
        for (var rootIndex = 0; rootIndex < roots.Length; rootIndex++)
        {
            var element = roots[rootIndex];
            if (HasOnlyAttributes(element, "id", "authorId", "status", "created") is false)
                throw Unsupported("a root comment contains task, reaction, or extension attributes");
            var children = element.Elements().ToArray();
            var anchors = children.Where(child => child.Name == Oac + "deMkLst" || child.Name == Oac + "txMkLst").ToArray();
            var position = SingleChild(element, P188 + "pos");
            var textBody = SingleChild(element, P188 + "txBody");
            var replyLists = children.Where(child => child.Name == P188 + "replyLst").ToArray();
            if (anchors.Length != 1 || position is null || textBody is null || replyLists.Length > 1 ||
                children.Any(child => child != anchors[0] && child != position && child != textBody && !replyLists.Contains(child)))
                throw Unsupported("a root comment has an unsupported anchor, child, or multiplicity");
            if (!HasOnlyAttributes(position, "x", "y") || position.HasElements)
                throw Unsupported("a comment position is not a plain x/y record");

            var rootComment = ParseComment(element, textBody, authors, ids, "root");
            var anchor = ParseAnchor(anchors[0], targets, nativeSlideId);
            var replies = new List<PresentationModernComment>();
            if (replyLists.Length == 1)
            {
                var replyList = replyLists[0];
                if (SignificantAttributes(replyList).Any() || replyList.Elements().Any(child => child.Name != P188 + "reply"))
                    throw Unsupported("the reply list contains unsupported content");
                foreach (var reply in replyList.Elements(P188 + "reply"))
                {
                    if (!HasOnlyAttributes(reply, "id", "authorId", "status", "created"))
                        throw Unsupported("a reply contains task, reaction, or extension attributes");
                    var replyText = SingleChild(reply, P188 + "txBody");
                    if (replyText is null || reply.Elements().Any(child => child != replyText))
                        throw Unsupported("a reply is not a single plain-text body");
                    replies.Add(ParseComment(reply, replyText, authors, ids, "reply"));
                }
            }
            if (!TryLong(position.Attribute("x")?.Value, out var x) || !TryLong(position.Attribute("y")?.Value, out var y))
                throw Unsupported("a comment position is invalid");

            var thread = new PresentationModernCommentThread
            {
                Id = rootComment.Id,
                TargetId = anchor.TargetId,
                Anchor = anchor.Wire,
                PositionXEmu = x,
                PositionYEmu = y,
                Root = rootComment,
            };
            thread.Replies.Add(replies);
            output.Add(thread);
        }
        return output;
    }

    private static PresentationModernComment ParseComment(
        XElement element,
        XElement textBody,
        IReadOnlyDictionary<string, PresentationModernComment> authors,
        ISet<string> ids,
        string kind)
    {
        var id = GuidAttribute(element, "id", $"modern comment {kind} ID");
        var authorId = GuidAttribute(element, "authorId", $"modern comment {kind} author ID");
        if (!ids.Add(id)) throw Unsupported($"duplicate modern comment ID {id}");
        if (!authors.TryGetValue(authorId, out var author)) throw Unsupported($"modern comment author {authorId} is missing");
        var status = element.Attribute("status")?.Value ?? string.Empty;
        if (!Statuses.Contains(status)) throw Unsupported($"modern comment status {status} is unsupported");
        var created = element.Attribute("created")?.Value ?? string.Empty;
        _ = ParseTimestamp(created, $"modern comment {id}");
        var text = ReadPlainText(textBody);
        return new PresentationModernComment
        {
            Id = id,
            AuthorId = authorId,
            Author = author.Author,
            Initials = author.Initials,
            UserId = author.UserId,
            ProviderId = author.ProviderId,
            Text = text,
            CreatedAt = created,
            Status = status,
        };
    }

    private static AnchorParse ParseAnchor(XElement element, IReadOnlyDictionary<uint, AnchorTarget> targets, uint nativeSlideId)
    {
        var kind = element.Name == Oac + "txMkLst"
            ? PresentationModernCommentAnchor.Types.Kind.TextRange
            : PresentationModernCommentAnchor.Types.Kind.Element;
        if (SignificantAttributes(element).Any()) throw Unsupported("a modern comment anchor has unsupported attributes");
        var slideList = SingleChild(element, Pc + "sldMkLst");
        if (slideList is null || SignificantAttributes(slideList).Any()) throw Unsupported("the slide moniker list is missing or extended");
        var slideChildren = slideList.Elements().ToArray();
        if (slideChildren.Length != 2 || slideChildren[0].Name != Pc + "docMk" || slideChildren[1].Name != Pc + "sldMk" ||
            SignificantAttributes(slideChildren[0]).Any() || !HasOnlyAttributes(slideChildren[1], "sldId") ||
            !TryUInt(slideChildren[1].Attribute("sldId")?.Value, out var anchoredSlideId) || anchoredSlideId != nativeSlideId)
            throw Unsupported("the modern comment slide moniker chain does not resolve to its source slide");

        var monikerElements = element.Elements().Where(child => Monikers.Contains(child.Name.LocalName)).ToArray();
        if (monikerElements.Length != 1 || monikerElements[0].Name.Namespace != Oac)
            throw Unsupported("the modern comment drawing moniker chain is nested or unsupported");
        var moniker = monikerElements[0];
        if (!HasOnlyAttributes(moniker, "id", "creationId") || moniker.HasElements ||
            !TryUInt(moniker.Attribute("id")?.Value, out var nativeId) || nativeId == 0 ||
            !targets.TryGetValue(nativeId, out var target) || target.Moniker != moniker.Name.LocalName)
            throw Unsupported("the modern comment drawing moniker does not resolve to a modeled element");
        var creationId = moniker.Attribute("creationId")?.Value?.ToUpperInvariant() ?? string.Empty;
        if (creationId.Length > 0 && (!IsGuid(creationId) || target.CreationId != creationId))
            throw Unsupported("the modern comment creationId does not match its drawing target");

        var anchor = new PresentationModernCommentAnchor { Kind = kind, NativeSlideId = nativeSlideId };
        anchor.Monikers.Add(new PresentationModernCommentMoniker { Type = moniker.Name.LocalName, NativeId = nativeId, CreationId = creationId });
        if (kind == PresentationModernCommentAnchor.Types.Kind.TextRange)
        {
            if (moniker.Name != Oac + "spMk" || !target.IsTextShape)
                throw Unsupported("a text-range anchor does not resolve to a text shape");
            var textMoniker = SingleChild(element, Oac + "txMk");
            if (textMoniker is null || !HasOnlyAttributes(textMoniker, "cp", "len") ||
                !TryUInt(textMoniker.Attribute("cp")?.Value, out var start) || !TryUInt(textMoniker.Attribute("len")?.Value, out var length) ||
                (ulong)start + length > (ulong)target.TextLength)
                throw Unsupported("the modern comment text range is invalid or out of bounds");
            var contexts = textMoniker.Elements().ToArray();
            if (contexts.Length > 1 || contexts.Any(child => child.Name != Oac + "context"))
                throw Unsupported("the modern comment text-range context is extended");
            anchor.TextStart = start;
            anchor.TextLength = length;
            if (contexts.Length == 1)
            {
                var context = contexts[0];
                if (!HasOnlyAttributes(context, "len", "hash") || !TryUInt(context.Attribute("hash")?.Value, out var hash))
                    throw Unsupported("the modern comment text-range context hash is invalid");
                anchor.ContextHash = hash;
                if (context.Attribute("len") is { } contextLength)
                {
                    if (!TryUInt(contextLength.Value, out var value)) throw Unsupported("the modern comment context length is invalid");
                    anchor.ContextLength = value;
                }
            }
            if (element.Elements().Any(child => child != slideList && child != moniker && child != textMoniker))
                throw Unsupported("the text-range anchor contains unsupported monikers");
            return new AnchorParse(anchor, $"{target.TargetId}/text");
        }

        if (element.Elements().Any(child => child != slideList && child != moniker))
            throw Unsupported("the drawing anchor contains unsupported monikers");
        return new AnchorParse(anchor, target.TargetId);
    }

    private static IReadOnlyDictionary<string, PresentationModernComment> ParseAuthors(byte[] bytes)
    {
        var document = LoadDocument(bytes);
        var root = document.Root;
        if (root?.Name != P188 + "authorLst" || SignificantAttributes(root).Any() || root.Elements().Any(child => child.Name != P188 + "author"))
            throw Unsupported("the modern author catalog root or children are outside the bounded profile");
        var authors = new Dictionary<string, PresentationModernComment>(StringComparer.OrdinalIgnoreCase);
        foreach (var author in root.Elements(P188 + "author"))
        {
            if (!HasOnlyAttributes(author, "id", "name", "initials", "userId", "providerId") || author.HasElements)
                throw Unsupported("a modern author contains unsupported metadata or extensions");
            var id = GuidAttribute(author, "id", "modern author ID");
            var name = RequiredMetadata(author, "name");
            var initials = OptionalMetadata(author, "initials");
            var userId = RequiredMetadata(author, "userId");
            var providerId = RequiredMetadata(author, "providerId");
            if (!authors.TryAdd(id, new PresentationModernComment
                {
                    AuthorId = id,
                    Author = name,
                    Initials = initials,
                    UserId = userId,
                    ProviderId = providerId,
                })) throw Unsupported($"duplicate modern author ID {id}");
        }
        if (authors.Count == 0 || authors.Count > MaxAuthorsPerPresentation)
            throw Unsupported("the modern author catalog is empty or exceeds its budget");
        return authors;
    }

    private static IReadOnlyDictionary<string, PresentationModernComment> CollectAuthors(IEnumerable<PresentationModernCommentThread> threads)
    {
        var authors = new Dictionary<string, PresentationModernComment>(StringComparer.OrdinalIgnoreCase);
        foreach (var comment in threads.SelectMany(thread => new[] { thread.Root }.Concat(thread.Replies)))
        {
            ValidateComment(comment, "source-free modern comment");
            if (authors.TryGetValue(comment.AuthorId, out var prior))
            {
                if (prior.Author != comment.Author || prior.Initials != comment.Initials || prior.UserId != comment.UserId || prior.ProviderId != comment.ProviderId)
                    throw Invalid($"Modern author {comment.AuthorId} has conflicting metadata.");
            }
            else authors.Add(comment.AuthorId, comment.Clone());
        }
        if (authors.Count > MaxAuthorsPerPresentation) throw Invalid("Presentation modern comments exceed the 256-author budget.");
        return authors;
    }

    private static void Validate(PresentationModernCommentThread thread, PresentationSlide slide, uint nativeSlideId, int slideIndex, bool sourceBound)
    {
        if (!IsGuid(thread.Id) || thread.Root is null || thread.Id != thread.Root.Id || string.IsNullOrWhiteSpace(thread.TargetId))
            throw Invalid($"Presentation slide {slideIndex + 1} has invalid modern-comment thread identity.");
        ValidateComment(thread.Root, $"Presentation slide {slideIndex + 1} modern root {thread.Id}");
        foreach (var reply in thread.Replies) ValidateComment(reply, $"Presentation slide {slideIndex + 1} modern reply {reply.Id}");
        var anchor = thread.Anchor ?? throw Invalid($"Presentation slide {slideIndex + 1} modern thread {thread.Id} has no anchor.");
        if (anchor.NativeSlideId != nativeSlideId || nativeSlideId < 256 || anchor.Monikers.Count != 1)
            throw Invalid($"Presentation slide {slideIndex + 1} modern thread {thread.Id} has an invalid slide or drawing moniker chain.");
        var moniker = anchor.Monikers[0];
        if (!Monikers.Contains(moniker.Type) || moniker.NativeId == 0 || (moniker.CreationId.Length > 0 && !IsGuid(moniker.CreationId)))
            throw Invalid($"Presentation slide {slideIndex + 1} modern thread {thread.Id} has an invalid drawing moniker.");
        if (anchor.Kind == PresentationModernCommentAnchor.Types.Kind.TextRange)
        {
            if (moniker.Type != "spMk" || !anchor.HasTextStart || !anchor.HasTextLength ||
                (ulong)anchor.TextStart + anchor.TextLength > int.MaxValue)
                throw Invalid($"Presentation slide {slideIndex + 1} modern thread {thread.Id} has an invalid text-range anchor.");
        }
        else if (anchor.Kind != PresentationModernCommentAnchor.Types.Kind.Element || anchor.HasTextStart || anchor.HasTextLength || anchor.HasContextHash || anchor.HasContextLength)
            throw Invalid($"Presentation slide {slideIndex + 1} modern thread {thread.Id} has an invalid element anchor.");
        if (!sourceBound)
        {
            var flattened = Flatten(slide.Elements).ToArray();
            if (moniker.NativeId < 2 || moniker.NativeId - 2 >= flattened.Length || flattened[checked((int)moniker.NativeId - 2)].Id != AnchorElementId(thread.TargetId))
                throw Invalid($"Presentation slide {slideIndex + 1} modern thread {thread.Id} does not resolve to its source-free target element.");
            var target = flattened[checked((int)moniker.NativeId - 2)];
            if (anchor.Kind == PresentationModernCommentAnchor.Types.Kind.TextRange &&
                (target.ContentCase != PresentationElement.ContentOneofCase.Shape ||
                 (ulong)anchor.TextStart + anchor.TextLength > (ulong)target.Shape.Text.Length))
                throw Invalid($"Presentation slide {slideIndex + 1} modern thread {thread.Id} text range exceeds its source-free shape text.");
        }
    }

    private static void ValidateComment(PresentationModernComment comment, string context)
    {
        if (!IsGuid(comment.Id) || !IsGuid(comment.AuthorId) || comment.Author.Length is 0 or > MaxMetadataLength ||
            comment.Initials.Length > MaxMetadataLength || comment.UserId.Length is 0 or > MaxMetadataLength ||
            comment.ProviderId.Length is 0 or > MaxMetadataLength || comment.Text.Length > MaxTextLength || !Statuses.Contains(comment.Status))
            throw Invalid($"{context} has invalid identity, author metadata, text, or status.");
        _ = ParseTimestamp(comment.CreatedAt, context);
    }

    private static XDocument AuthorsDocument(IReadOnlyDictionary<string, PresentationModernComment> authors) => new(
        new XDeclaration("1.0", "UTF-8", "yes"),
        new XElement(P188 + "authorLst",
            new XAttribute(XNamespace.Xmlns + "p188", P188),
            authors.Values.Select(author => new XElement(P188 + "author",
                new XAttribute("id", author.AuthorId),
                new XAttribute("name", author.Author),
                new XAttribute("initials", author.Initials),
                new XAttribute("userId", author.UserId),
                new XAttribute("providerId", author.ProviderId)))));

    private static XDocument CommentsDocument(IEnumerable<PresentationModernCommentThread> threads) => new(
        new XDeclaration("1.0", "UTF-8", "yes"),
        new XElement(P188 + "cmLst",
            new XAttribute(XNamespace.Xmlns + "p188", P188),
            new XAttribute(XNamespace.Xmlns + "a", Drawing),
            new XAttribute(XNamespace.Xmlns + "oac", Oac),
            new XAttribute(XNamespace.Xmlns + "pc", Pc),
            threads.Select(CommentElement)));

    private static XElement CommentElement(PresentationModernCommentThread thread)
    {
        var children = new List<object>
        {
            AnchorElement(thread.Anchor),
            new XElement(P188 + "pos", new XAttribute("x", thread.PositionXEmu), new XAttribute("y", thread.PositionYEmu)),
        };
        if (thread.Replies.Count > 0)
            children.Add(new XElement(P188 + "replyLst", thread.Replies.Select(reply => CommentNode("reply", reply))));
        children.Add(TextBody(thread.Root.Text));
        return new XElement(P188 + "cm", CommentAttributes(thread.Root), children);
    }

    private static XElement CommentNode(string localName, PresentationModernComment comment) =>
        new(P188 + localName, CommentAttributes(comment), TextBody(comment.Text));

    private static IEnumerable<XAttribute> CommentAttributes(PresentationModernComment comment)
    {
        yield return new XAttribute("id", comment.Id);
        yield return new XAttribute("authorId", comment.AuthorId);
        yield return new XAttribute("status", comment.Status);
        yield return new XAttribute("created", DateTimeOffset.Parse(comment.CreatedAt, CultureInfo.InvariantCulture).ToUniversalTime().ToString("O", CultureInfo.InvariantCulture));
    }

    private static XElement AnchorElement(PresentationModernCommentAnchor anchor)
    {
        var moniker = anchor.Monikers.Single();
        var monikerElement = new XElement(Oac + moniker.Type, new XAttribute("id", moniker.NativeId));
        if (!string.IsNullOrWhiteSpace(moniker.CreationId)) monikerElement.Add(new XAttribute("creationId", moniker.CreationId));
        var slideList = new XElement(Pc + "sldMkLst", new XElement(Pc + "docMk"), new XElement(Pc + "sldMk", new XAttribute("sldId", anchor.NativeSlideId)));
        if (anchor.Kind == PresentationModernCommentAnchor.Types.Kind.Element)
            return new XElement(Oac + "deMkLst", slideList, monikerElement);
        var textMoniker = new XElement(Oac + "txMk", new XAttribute("cp", anchor.TextStart), new XAttribute("len", anchor.TextLength));
        if (anchor.HasContextHash)
        {
            var context = new XElement(Oac + "context", new XAttribute("hash", anchor.ContextHash));
            if (anchor.HasContextLength) context.Add(new XAttribute("len", anchor.ContextLength));
            textMoniker.Add(context);
        }
        return new XElement(Oac + "txMkLst", slideList, monikerElement, textMoniker);
    }

    private static XElement TextBody(string text)
    {
        var value = new XElement(Drawing + "t", text);
        SetSpace(value, text);
        return new XElement(P188 + "txBody",
            new XElement(Drawing + "bodyPr"),
            new XElement(Drawing + "lstStyle"),
            new XElement(Drawing + "p", new XElement(Drawing + "r", value)));
    }

    private static string ReadPlainText(XElement textBody)
    {
        var children = textBody.Elements().ToArray();
        if (SignificantAttributes(textBody).Any() || children.Length != 3 || children[0].Name != Drawing + "bodyPr" || children[1].Name != Drawing + "lstStyle" || children[2].Name != Drawing + "p" ||
            SignificantAttributes(children[0]).Any() || children[0].HasElements || SignificantAttributes(children[1]).Any() || children[1].HasElements)
            throw Unsupported("a modern comment text body is not the bounded plain-text profile");
        var paragraph = children[2];
        if (SignificantAttributes(paragraph).Any()) throw Unsupported("a modern comment paragraph has formatting");
        var runs = paragraph.Elements().ToArray();
        if (runs.Length != 1 || runs[0].Name != Drawing + "r" || SignificantAttributes(runs[0]).Any())
            throw Unsupported("a modern comment paragraph is not a single plain run");
        var runChildren = runs[0].Elements().ToArray();
        if (runChildren.Length != 1 || runChildren[0].Name != Drawing + "t" || runChildren[0].HasElements ||
            runChildren[0].Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != Xml + "space"))
            throw Unsupported("a modern comment run is not plain text");
        var text = runChildren[0].Value;
        if (text.Length > MaxTextLength) throw Unsupported("a modern comment exceeds the text budget");
        return text;
    }

    private static bool ApplyMutableComment(XElement element, PresentationModernComment original, PresentationModernComment requested)
    {
        var changed = false;
        if (original.Status != requested.Status)
        {
            element.SetAttributeValue("status", requested.Status);
            changed = true;
        }
        if (original.Text != requested.Text)
        {
            var text = element.Element(P188 + "txBody")!.Descendants(Drawing + "t").Single();
            text.Value = requested.Text;
            SetSpace(text, requested.Text);
            changed = true;
        }
        return changed;
    }

    private static void SetSpace(XElement text, string value)
    {
        if (value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1]))) text.SetAttributeValue(Xml + "space", "preserve");
        else text.SetAttributeValue(Xml + "space", null);
    }

    private static IReadOnlyDictionary<uint, AnchorTarget> AnchorTargets(IReadOnlyList<OpenXmlElement> elements, IReadOnlyDictionary<uint, string> ids)
    {
        var output = new Dictionary<uint, AnchorTarget>();
        CollectTargets(elements, ids, output);
        return output;
    }

    private static void CollectTargets(IReadOnlyList<OpenXmlElement> elements, IReadOnlyDictionary<uint, string> ids, IDictionary<uint, AnchorTarget> output)
    {
        foreach (var element in elements)
        {
            var nativeId = element.Descendants<P.NonVisualDrawingProperties>().FirstOrDefault()?.Id?.Value;
            if (nativeId is not uint id || !ids.TryGetValue(id, out var targetId)) continue;
            var moniker = element switch
            {
                P.Shape => "spMk",
                P.GraphicFrame => "graphicFrameMk",
                P.ConnectionShape => "cxnSpMk",
                P.Picture => "picMk",
                P.GroupShape => "grpSpMk",
                _ => string.Empty,
            };
            if (moniker.Length == 0) continue;
            var creationIds = element.Descendants<A16.CreationId>().Select(value => value.Id?.Value?.ToUpperInvariant()).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            if (creationIds.Length > 1) continue;
            var text = element is P.Shape shape ? PlainShapeText(shape) : string.Empty;
            output[id] = new AnchorTarget(targetId, moniker, creationIds.SingleOrDefault() ?? string.Empty, element is P.Shape, text.Length);
        }
    }

    private static string PlainShapeText(P.Shape shape) => string.Join("\r", shape.TextBody?.Elements<A.Paragraph>().Select(paragraph =>
        string.Concat(paragraph.ChildElements.Select(child => child switch
        {
            A.Run run => run.Text?.Text ?? string.Empty,
            A.Field field => field.Text?.Text ?? string.Empty,
            A.Break => "\v",
            _ => string.Empty,
        }))) ?? []);

    private static IEnumerable<PresentationElement> Flatten(IEnumerable<PresentationElement> elements)
    {
        foreach (var element in elements)
        {
            yield return element;
            if (element.ContentCase == PresentationElement.ContentOneofCase.Group)
                foreach (var child in Flatten(element.Group.Children)) yield return child;
        }
    }

    private static string AnchorElementId(string targetId) => targetId.EndsWith("/text", StringComparison.Ordinal) ? targetId[..^5] : targetId;

    private static string FixedTopologyHash(PresentationModernCommentThread thread)
    {
        var builder = new StringBuilder();
        builder.Append(thread.Id).Append('\0').Append(thread.TargetId).Append('\0').Append(thread.PositionXEmu).Append('\0').Append(thread.PositionYEmu).Append('\0');
        var anchor = thread.Anchor;
        builder.Append((int)(anchor?.Kind ?? 0)).Append('\0').Append(anchor?.NativeSlideId ?? 0).Append('\0');
        if (anchor is not null)
        {
            foreach (var moniker in anchor.Monikers) builder.Append(moniker.Type).Append('\0').Append(moniker.NativeId).Append('\0').Append(moniker.CreationId).Append('\0');
            builder.Append(anchor.HasTextStart ? anchor.TextStart : "-").Append('\0')
                .Append(anchor.HasTextLength ? anchor.TextLength : "-").Append('\0')
                .Append(anchor.HasContextLength ? anchor.ContextLength : "-").Append('\0')
                .Append(anchor.HasContextHash ? anchor.ContextHash : "-").Append('\0');
        }
        AppendFixedComment(builder, thread.Root);
        foreach (var reply in thread.Replies) AppendFixedComment(builder, reply);
        return Hash(Encoding.UTF8.GetBytes(builder.ToString()));
    }

    private static void AppendFixedComment(StringBuilder builder, PresentationModernComment comment) => builder
        .Append(comment.Id).Append('\0').Append(comment.AuthorId).Append('\0').Append(comment.Author).Append('\0')
        .Append(comment.Initials).Append('\0').Append(comment.UserId).Append('\0').Append(comment.ProviderId).Append('\0')
        .Append(comment.CreatedAt).Append('\0');

    private static PptxModernCommentsChange? RejectModeledEdit(PresentationSlide requested, OpenXmlPart part, int slideIndex, string reason)
    {
        if (requested.ModernComments.Count == 0) return null;
        throw new CodecException("unsupported_presentation_comment_edit", $"Presentation slide {slideIndex + 1} modern comments are outside the bounded editable profile: {reason}.", PartPath(part));
    }

    private static IReadOnlyList<PresentationModernCommentThread> Unsupported(OpenXmlPart part, int slideIndex, IList<Diagnostic> diagnostics, string reason)
    {
        diagnostics.Add(CodecProtocol.Warning(
            "unsupported_presentation_modern_comments_preserved",
            $"Presentation slide {slideIndex + 1} has modern comments outside the bounded profile ({reason}); the graph remains opaque and source-bound.",
            PartPath(part)));
        return [];
    }

    private static CodecException TopologyChanged(int slideIndex, OpenXmlPart part) => new(
        "presentation_comment_topology_changed",
        $"Presentation slide {slideIndex + 1} changed modern-comment identity, author/date metadata, anchor, position, topology, or source binding. Only text and status are editable.",
        PartPath(part));

    private static XElement? SingleChild(XElement owner, XName name)
    {
        var children = owner.Elements(name).ToArray();
        return children.Length == 1 ? children[0] : null;
    }

    private static IEnumerable<XAttribute> SignificantAttributes(XElement element) => element.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration);
    private static bool HasOnlyAttributes(XElement element, params string[] names)
    {
        var allowed = new HashSet<string>(names, StringComparer.Ordinal);
        return SignificantAttributes(element).All(attribute => attribute.Name.Namespace == XNamespace.None && allowed.Contains(attribute.Name.LocalName));
    }

    private static string GuidAttribute(XElement element, string name, string label)
    {
        var value = element.Attribute(name)?.Value?.ToUpperInvariant() ?? string.Empty;
        if (!IsGuid(value)) throw Unsupported($"{label} is missing or invalid");
        return value;
    }

    private static string RequiredMetadata(XElement element, string name)
    {
        var value = element.Attribute(name)?.Value ?? string.Empty;
        if (value.Length is 0 or > MaxMetadataLength) throw Unsupported($"modern author {name} is missing or too long");
        return value;
    }

    private static string OptionalMetadata(XElement element, string name)
    {
        var value = element.Attribute(name)?.Value ?? string.Empty;
        if (value.Length > MaxMetadataLength) throw Unsupported($"modern author {name} is too long");
        return value;
    }

    private static bool IsGuid(string value) => value.Length == 38 && value[0] == '{' && value[^1] == '}' && Guid.TryParse(value[1..^1], out _);
    private static DateTimeOffset ParseTimestamp(string value, string context) => DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var timestamp)
        ? timestamp
        : throw Unsupported($"{context} has an invalid date-time");
    private static bool TryUInt(string? value, out uint result) => uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out result);
    private static bool TryLong(string? value, out long result) => long.TryParse(value, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out result);

    private static T[] Parts<T>(OpenXmlPartContainer owner) where T : OpenXmlPart => owner.Parts.Select(pair => pair.OpenXmlPart).OfType<T>().ToArray();
    private static bool HasRelationships(OpenXmlPart part) => part.Parts.Any() || part.ExternalRelationships.Any() || part.HyperlinkRelationships.Any() || part.DataPartReferenceRelationships.Any();
    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static byte[] ReadBytes(OpenXmlPart part)
    {
        using var source = part.GetStream(FileMode.Open, FileAccess.Read);
        using var output = new MemoryStream();
        source.CopyTo(output);
        return output.ToArray();
    }

    private static void WriteBytes(OpenXmlPart part, byte[] bytes)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        stream.Write(bytes);
    }

    private static XDocument LoadDocument(byte[] bytes)
    {
        try
        {
            using var stream = new MemoryStream(bytes, writable: false);
            return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        }
        catch (Exception exception) when (exception is XmlException or InvalidOperationException)
        {
            throw Unsupported("the modern comment XML is malformed");
        }
    }

    private static byte[] DocumentBytes(XDocument document)
    {
        using var output = new MemoryStream();
        using (var writer = XmlWriter.Create(output, new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(false),
            Indent = false,
            OmitXmlDeclaration = false,
            CloseOutput = false,
        })) document.Save(writer);
        return output.ToArray();
    }

    private static void WriteDocument(OpenXmlPart part, XDocument document) => WriteBytes(part, DocumentBytes(document));
    private static CodecException Invalid(string message) => new("invalid_presentation_modern_comment", message);
    private static UnsupportedProfileException Unsupported(string message) => new(message);

    private sealed record AnchorTarget(string TargetId, string Moniker, string CreationId, bool IsTextShape, int TextLength);
    private sealed record AnchorParse(PresentationModernCommentAnchor Wire, string TargetId);
    private sealed class UnsupportedProfileException(string message) : Exception(message);
}
