using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

internal sealed record DocxClassicCommentSource(
    W.Comment Element,
    W.Paragraph Target,
    W.CommentRangeStart Start,
    W.CommentRangeEnd End,
    W.Run ReferenceRun,
    DocumentComment Artifact);

internal sealed record DocxClassicCommentGraph(
    WordprocessingCommentsPart Part,
    IReadOnlyList<DocxClassicCommentSource> Comments,
    DocxExtendedCommentGraph Extended);

// Owns classic w:comment bodies and their document-story anchors. The optional
// Office 2013+ thread/support parts are delegated to DocxExtendedCommentCodec.
// Roots use one exact whole-paragraph anchor; direct replies share that root
// anchor and never pretend to own an independent story range.
internal static class DocxClassicCommentCodec
{
    private const string WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const int MaxAuthorLength = 255;
    private const int MaxProviderIdLength = 100;
    private const int MaxUserIdLength = 300;
    private const int MaxInitialsLength = 9;
    private const int MaxTimestampLength = 64;
    private const int MaxTextLength = 1_000_000;

    internal static void Read(
        DocxPartContext context,
        W.Body body,
        DocumentArtifact document,
        ref ulong semanticItems,
        EffectiveCodecLimits limits,
        ICollection<Diagnostic> diagnostics)
    {
        if (context.Owner.WordprocessingCommentsPart is null) return;
        if (!TryReadGraph(context, body, document, out var graph, out var reason))
        {
            diagnostics.Add(CodecProtocol.Warning(
                "unsupported_document_comments_preserved",
                $"Preserved a DOCX comment graph that is outside the editable bounded-comment profile: {reason}",
                "word/comments.xml"));
            return;
        }

        foreach (var source in graph.Comments)
        {
            document.Comments.Add(source.Artifact);
            semanticItems++;
            if (semanticItems > limits.MaxCells)
                throw new CodecException(
                    "document_item_budget_exceeded",
                    $"DOCX document exceeds max_cells semantic-item budget ({limits.MaxCells}).",
                    "word/comments.xml");
        }
    }

    internal static void Validate(DocumentArtifact document, EffectiveCodecLimits limits)
    {
        if (document.Comments.Count == 0) return;

        var blocks = new Dictionary<string, DocumentBlock>(StringComparer.Ordinal);
        foreach (var block in document.Blocks)
        {
            if (string.IsNullOrWhiteSpace(block.Id) || !blocks.TryAdd(block.Id, block))
                throw new CodecException("invalid_document_comment", "Document comments require unique, non-empty target block IDs.");
        }

        if ((ulong)document.Comments.Count > limits.MaxCells)
            throw new CodecException(
                "document_item_budget_exceeded",
                $"Document has {document.Comments.Count} comments and exceeds max_cells ({limits.MaxCells}).");

        var ids = new HashSet<string>(StringComparer.Ordinal);
        var commentsById = new Dictionary<string, DocumentComment>(StringComparer.Ordinal);
        foreach (var comment in document.Comments)
        {
            if (string.IsNullOrWhiteSpace(comment.Id) || !ids.Add(comment.Id))
                throw new CodecException("invalid_document_comment", "Document comments require unique, non-empty IDs.");
            commentsById.Add(comment.Id, comment);
            if (!blocks.TryGetValue(comment.TargetBlockId, out var target))
                throw new CodecException(
                    "invalid_document_comment",
                    $"Document comment {comment.Id} targets missing block {comment.TargetBlockId}.");
            if (target.ContentCase is not DocumentBlock.ContentOneofCase.Paragraph and
                not DocumentBlock.ContentOneofCase.Hyperlink and
                not DocumentBlock.ContentOneofCase.Field)
                throw new CodecException(
                    "unsupported_document_comment_target",
                    $"Document comment {comment.Id} requires a top-level paragraph, hyperlink, or field target; {comment.TargetBlockId} is {target.ContentCase}.");

            ValidateText(comment.Author, $"Document comment {comment.Id} author", 1, MaxAuthorLength);
            ValidateText(comment.Text, $"Document comment {comment.Id} text", 0, MaxTextLength);
            if (comment.HasInitials)
                ValidateText(comment.Initials, $"Document comment {comment.Id} initials", 1, MaxInitialsLength);
            if (comment.HasCreatedAt)
            {
                ValidateText(comment.CreatedAt, $"Document comment {comment.Id} created_at", 1, MaxTimestampLength);
                if (!DateTimeOffset.TryParse(
                        comment.CreatedAt,
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.RoundtripKind,
                        out _))
                    throw new CodecException(
                        "invalid_document_comment",
                        $"Document comment {comment.Id} created_at must be an ISO 8601 date-time.");
            }
            ValidateOptionalParagraphId(comment.ParagraphId, $"Document comment {comment.Id} paragraph_id");
            ValidateOptionalDurableId(comment.DurableId, $"Document comment {comment.Id} durable_id");
            if (comment.HasDateUtc)
            {
                ValidateText(comment.DateUtc, $"Document comment {comment.Id} date_utc", 1, MaxTimestampLength);
                if (!DateTimeOffset.TryParse(comment.DateUtc, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out _))
                    throw new CodecException("invalid_document_comment", $"Document comment {comment.Id} date_utc must be an ISO 8601 date-time.");
            }
            if (comment.Person is not null)
            {
                ValidateText(comment.Person.ProviderId, $"Document comment {comment.Id} person provider_id", 1, MaxProviderIdLength);
                ValidateText(comment.Person.UserId, $"Document comment {comment.Id} person user_id", 1, MaxUserIdLength);
            }

            if (comment.Source is { } source)
            {
                if (string.IsNullOrWhiteSpace(source.NativeCommentId) ||
                    string.IsNullOrWhiteSpace(source.CommentElementSha256) ||
                    string.IsNullOrWhiteSpace(source.SemanticSha256) ||
                    string.IsNullOrWhiteSpace(source.ResidualSha256) ||
                    string.IsNullOrWhiteSpace(source.AnchorSha256))
                    throw new CodecException(
                        "invalid_document_comment_source_binding",
                        $"Document comment {comment.Id} has an incomplete source binding.");
                if (source.ThreadEditable && string.IsNullOrWhiteSpace(source.ExtendedGraphSha256))
                    throw new CodecException(
                        "invalid_document_comment_source_binding",
                        $"Modern document comment {comment.Id} has no extended graph binding.");
            }
        }

        foreach (var comment in document.Comments.Where(comment => comment.ParentCommentId.Length > 0))
        {
            if (!commentsById.TryGetValue(comment.ParentCommentId, out var parent))
                throw new CodecException("invalid_document_comment_thread", $"Document comment {comment.Id} references missing parent {comment.ParentCommentId}.");
            if (parent.ParentCommentId.Length > 0)
                throw new CodecException("unsupported_document_comment_thread", $"Document comment {comment.Id} is nested; only direct replies to roots are supported.");
            if (!parent.TargetBlockId.Equals(comment.TargetBlockId, StringComparison.Ordinal))
                throw new CodecException("invalid_document_comment_thread", $"Document comment {comment.Id} and root {parent.Id} must target the same block.");
            if (comment.HasIntelligentPlaceholder && comment.IntelligentPlaceholder)
                throw new CodecException("invalid_document_comment_thread", $"Document reply {comment.Id} cannot be an intelligent placeholder.");
        }
        foreach (var authorGroup in document.Comments.GroupBy(comment => comment.Author, StringComparer.Ordinal))
        {
            var profiles = authorGroup.Select(comment => comment.Person is null
                    ? string.Empty
                    : Convert.ToBase64String(comment.Person.ToByteArray()))
                .Distinct(StringComparer.Ordinal)
                .Count();
            if (profiles != 1)
                throw new CodecException("invalid_document_comment", $"Document comment author {authorGroup.Key} has conflicting people metadata.");
        }
    }

    internal static void Author(DocxPartContext context, W.Body body, DocumentArtifact document)
    {
        if (document.Comments.Count == 0) return;
        var part = context.Owner.AddNewPart<WordprocessingCommentsPart>();
        part.Comments = new W.Comments();

        var targets = new Dictionary<string, W.Paragraph>(StringComparer.Ordinal);
        for (var index = 0; index < document.Blocks.Count; index++)
        {
            if (index >= body.ChildElements.Count || body.ChildElements[index] is not W.Paragraph paragraph) continue;
            targets[document.Blocks[index].Id] = paragraph;
        }

        var authored = new List<(DocumentComment Artifact, string NativeId, W.Paragraph Target, W.Comment Element)>();
        for (var index = 0; index < document.Comments.Count; index++)
        {
            var artifact = document.Comments[index];
            if (!targets.TryGetValue(artifact.TargetBlockId, out var target))
                throw new CodecException(
                    "unsupported_document_comment_target",
                    $"Document comment {artifact.Id} does not resolve to an authored top-level paragraph.");
            var nativeId = index.ToString(CultureInfo.InvariantCulture);
            var element = BuildComment(artifact, nativeId);
            part.Comments.Append(element);
            authored.Add((artifact, nativeId, target, element));
        }

        foreach (var group in authored.Where(item => item.Artifact.ParentCommentId.Length == 0).GroupBy(item => item.Target))
            AddWholeParagraphAnchors(group.Key, group.Select(item => (item.Artifact, item.NativeId, item.Target)).ToArray());
        DocxExtendedCommentCodec.Author(context.Owner, authored.Select(item => (item.Artifact, item.NativeId, item.Element)).ToArray());
        part.Comments.Save();
    }

    internal static void AssertModeledCommentsWereNotRemoved(byte[] sourceBytes, DocumentArtifact document)
    {
        if (document.Comments.Count != 0) return;
        using var stream = new MemoryStream(sourceBytes, writable: false);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var mainPart = package.MainDocumentPart;
        var body = mainPart?.Document?.Body;
        if (mainPart?.WordprocessingCommentsPart is null || body is null) return;
        var context = new DocxPartContext(mainPart);
        if (TryReadGraph(context, body, document, out var graph, out _) && graph.Comments.Count > 0)
            throw new CodecException(
                "document_comment_topology_changed",
                $"Source-preserving DOCX export requires the original {graph.Comments.Count}-comment topology; the artifact contains 0 comments.",
                "word/comments.xml");
    }

    internal static void ApplySource(
        DocxPartContext context,
        W.Body body,
        DocumentArtifact document)
    {
        // Unsupported comment parts are opaque. Do not materialize their SDK
        // DOM in an editable package, because AutoSave may normalize bytes that
        // the caller did not ask this codec to change.
        if (document.Comments.Count == 0) return;
        var part = context.Owner.WordprocessingCommentsPart;
        if (part is null)
        {
            throw new CodecException(
                "document_comment_topology_changed",
                "Source-preserving DOCX export cannot add a Comments part in the current fixed-topology slice.",
                "word/document.xml");
        }

        if (!TryReadGraph(context, body, document, out var graph, out var reason))
        {
            throw new CodecException(
                "unsupported_document_comment_edit",
                $"The source DOCX comment graph is preserved but not editable: {reason}",
                "word/comments.xml");
        }
        if (graph.Comments.Count != document.Comments.Count)
            throw new CodecException(
                "document_comment_topology_changed",
                $"Source-preserving DOCX export requires the original {graph.Comments.Count}-comment topology; the artifact contains {document.Comments.Count} comments.",
                "word/comments.xml");

        var sourceByNativeId = graph.Comments.ToDictionary(
            item => item.Artifact.Source.NativeCommentId,
            StringComparer.Ordinal);
        var classicChanged = false;
        var extendedChanged = false;
        foreach (var requested in document.Comments)
        {
            var binding = requested.Source ?? throw new CodecException(
                "missing_document_comment_source_binding",
                $"Document comment {requested.Id} is missing its source binding.",
                "word/comments.xml");
            if (!sourceByNativeId.TryGetValue(binding.NativeCommentId, out var source))
                throw new CodecException(
                    "document_comment_source_binding_mismatch",
                    $"Document comment {requested.Id} does not resolve to its native source comment.",
                    "word/comments.xml");
            var original = source.Artifact;
            AssertBindingMatches(requested, binding, graph, source, original);

            var requestedHash = SemanticHash(requested);
            if (requestedHash.Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase)) continue;
            if (!binding.Editable || requested.Id != original.Id || requested.TargetBlockId != original.TargetBlockId ||
                requested.ParentCommentId != original.ParentCommentId || requested.ParagraphId != original.ParagraphId ||
                requested.DurableId != original.DurableId || requested.HasDateUtc != original.HasDateUtc ||
                requested.DateUtc != original.DateUtc || requested.HasIntelligentPlaceholder != original.HasIntelligentPlaceholder ||
                requested.IntelligentPlaceholder != original.IntelligentPlaceholder ||
                !Equals(requested.Person, original.Person) ||
                graph.Extended.PeoplePart is not null && requested.Author != original.Author)
                throw new CodecException(
                    "unsupported_document_comment_edit",
                    $"Document comment {requested.Id} changes source-bound identity, thread topology, durable/person metadata, or a people-bound author.",
                    "word/comments.xml");

            if (!ClassicValuesEqual(requested, original))
            {
                ApplyValues(source.Element, requested);
                classicChanged = true;
            }
            extendedChanged |= DocxExtendedCommentCodec.ApplyResolved(graph.Extended, binding.NativeCommentId, requested);
            if (!GraphResidualHash(graph.Part.Comments!).Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "document_comment_residual_not_preserved",
                    $"Document comment {requested.Id} changed unmodeled comment formatting or extension markup.",
                    "word/comments.xml");
            if (!AnchorHash(source.Start, source.End, source.ReferenceRun).Equals(binding.AnchorSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "document_comment_anchor_not_preserved",
                    $"Document comment {requested.Id} changed its source anchor triplet.",
                    "word/document.xml");
        }

        if (!classicChanged && !extendedChanged) return;
        if (classicChanged)
        {
            graph.Part.Comments!.Save();
            context.MarkCommentsMutated(graph.Part);
        }
        if (extendedChanged) DocxExtendedCommentCodec.SaveResolved(graph.Extended, context);

        if (!TryReadGraph(context, body, document, out var verifiedGraph, out var verifyReason))
            throw new CodecException(
                "document_comment_semantics_not_applied",
                $"Edited document comments no longer match the bounded profile: {verifyReason}",
                "word/comments.xml");
        var verifiedByNativeId = verifiedGraph.Comments.ToDictionary(item => item.Artifact.Source.NativeCommentId, StringComparer.Ordinal);
        foreach (var requested in document.Comments)
        {
            var nativeId = requested.Source!.NativeCommentId;
            if (!verifiedByNativeId.TryGetValue(nativeId, out var verified) ||
                !SemanticHash(verified.Artifact).Equals(SemanticHash(requested), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "document_comment_semantics_not_applied",
                    $"Document comment {requested.Id} does not match the requested semantics after editing.",
                    "word/comments.xml");
        }
    }

    private static void AssertBindingMatches(
        DocumentComment requested,
        DocumentCommentSourceBinding binding,
        DocxClassicCommentGraph graph,
        DocxClassicCommentSource source,
        DocumentComment original)
    {
        var actual = original.Source;
        if (binding.NativeCommentId != actual.NativeCommentId ||
            binding.TargetBodyIndex != actual.TargetBodyIndex ||
            !binding.CommentElementSha256.Equals(actual.CommentElementSha256, StringComparison.OrdinalIgnoreCase) ||
            !binding.SemanticSha256.Equals(actual.SemanticSha256, StringComparison.OrdinalIgnoreCase) ||
            !binding.ResidualSha256.Equals(actual.ResidualSha256, StringComparison.OrdinalIgnoreCase) ||
            !binding.AnchorSha256.Equals(actual.AnchorSha256, StringComparison.OrdinalIgnoreCase) ||
            binding.Editable != actual.Editable ||
            !binding.ExtendedGraphSha256.Equals(actual.ExtendedGraphSha256, StringComparison.OrdinalIgnoreCase) ||
            !binding.CommentsIdsGraphSha256.Equals(actual.CommentsIdsGraphSha256, StringComparison.OrdinalIgnoreCase) ||
            !binding.CommentsExtensibleGraphSha256.Equals(actual.CommentsExtensibleGraphSha256, StringComparison.OrdinalIgnoreCase) ||
            !binding.PeopleGraphSha256.Equals(actual.PeopleGraphSha256, StringComparison.OrdinalIgnoreCase) ||
            binding.ThreadEditable != actual.ThreadEditable ||
            !HashElement(source.Element).Equals(binding.CommentElementSha256, StringComparison.OrdinalIgnoreCase) ||
            !GraphResidualHash(graph.Part.Comments!).Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase) ||
            !graph.Extended.ExtendedGraphSha256.Equals(binding.ExtendedGraphSha256, StringComparison.OrdinalIgnoreCase) ||
            !graph.Extended.CommentsIdsGraphSha256.Equals(binding.CommentsIdsGraphSha256, StringComparison.OrdinalIgnoreCase) ||
            !graph.Extended.CommentsExtensibleGraphSha256.Equals(binding.CommentsExtensibleGraphSha256, StringComparison.OrdinalIgnoreCase) ||
            !graph.Extended.PeopleGraphSha256.Equals(binding.PeopleGraphSha256, StringComparison.OrdinalIgnoreCase) ||
            !AnchorHash(source.Start, source.End, source.ReferenceRun).Equals(binding.AnchorSha256, StringComparison.OrdinalIgnoreCase) ||
            !SemanticHash(original).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "document_comment_source_binding_mismatch",
                $"Document comment {requested.Id} no longer matches its source integrity binding.",
                "word/comments.xml");
    }

    private static bool TryReadGraph(
        DocxPartContext context,
        W.Body body,
        DocumentArtifact document,
        out DocxClassicCommentGraph graph,
        out string reason)
    {
        graph = null!;
        reason = string.Empty;
        var part = context.Owner.WordprocessingCommentsPart;
        if (part?.Comments is null)
        {
            reason = "the Comments part has no w:comments root";
            return false;
        }
        if (part.Comments.ChildElements.Any(child => child is not W.Comment))
        {
            reason = "w:comments contains an unsupported child";
            return false;
        }

        var elements = part.Comments.Elements<W.Comment>().ToArray();
        var nativeIds = new HashSet<string>(StringComparer.Ordinal);
        var numericIds = new HashSet<int>();
        foreach (var element in elements)
        {
            var nativeId = element.Id?.Value;
            if (string.IsNullOrWhiteSpace(nativeId) ||
                !int.TryParse(nativeId, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var numericId) ||
                numericId < 0 || !nativeIds.Add(nativeId) || !numericIds.Add(numericId))
            {
                reason = "comment IDs must be unique non-negative decimal integers";
                return false;
            }
            if (!IsSimpleCommentBody(element))
            {
                reason = $"comment {nativeId} does not contain one paragraph/run/text body";
                return false;
            }
        }
        if (!DocxExtendedCommentCodec.TryRead(context.Owner, elements, out var extended, out reason)) return false;

        var modelIdByNativeId = elements.Select((element, index) => (NativeId: element.Id!.Value!, ModelId: $"document/comment/{index + 1}"))
            .ToDictionary(item => item.NativeId, item => item.ModelId, StringComparer.Ordinal);
        var nativeIdByParagraphId = extended.ByNativeCommentId
            .ToDictionary(item => item.Value.ParagraphId, item => item.Key, StringComparer.Ordinal);
        var rootNativeIds = elements.Select(element => element.Id!.Value!)
            .Where(nativeId => !extended.IsModern || extended.ByNativeCommentId[nativeId].ParentParagraphId.Length == 0)
            .ToHashSet(StringComparer.Ordinal);

        var starts = body.Descendants<W.CommentRangeStart>().ToArray();
        var ends = body.Descendants<W.CommentRangeEnd>().ToArray();
        var references = body.Descendants<W.CommentReference>().ToArray();
        if (starts.Length != rootNativeIds.Count || ends.Length != rootNativeIds.Count || references.Length != rootNativeIds.Count)
        {
            reason = "root-comment/start/end/reference counts do not match";
            return false;
        }
        if (starts.Any(item => !rootNativeIds.Contains(item.Id?.Value ?? string.Empty)) ||
            ends.Any(item => !rootNativeIds.Contains(item.Id?.Value ?? string.Empty)) ||
            references.Any(item => !rootNativeIds.Contains(item.Id?.Value ?? string.Empty)))
        {
            reason = "the body contains a dangling or foreign comment anchor";
            return false;
        }

        var blockByBodyIndex = new Dictionary<uint, DocumentBlock>();
        foreach (var block in document.Blocks.Where(block => block.Source is not null))
        {
            if (!blockByBodyIndex.TryAdd(block.Source.BodyIndex, block))
            {
                reason = $"multiple modeled blocks claim body index {block.Source.BodyIndex}";
                return false;
            }
        }
        var rootAnchors = new Dictionary<string, (W.Paragraph Target, W.CommentRangeStart Start, W.CommentRangeEnd End, W.Run ReferenceRun, uint BodyIndex, DocumentBlock Block)>(StringComparer.Ordinal);
        foreach (var nativeId in rootNativeIds)
        {
            var matchingStarts = starts.Where(item => item.Id?.Value == nativeId).ToArray();
            var matchingEnds = ends.Where(item => item.Id?.Value == nativeId).ToArray();
            var matchingReferences = references.Where(item => item.Id?.Value == nativeId).ToArray();
            if (matchingStarts.Length != 1 || matchingEnds.Length != 1 || matchingReferences.Length != 1 ||
                matchingStarts[0].Parent is not W.Paragraph target ||
                matchingEnds[0].Parent is not W.Paragraph endTarget ||
                matchingReferences[0].Parent is not W.Run referenceRun ||
                referenceRun.Parent is not W.Paragraph referenceTarget ||
                !ReferenceEquals(target, endTarget) || !ReferenceEquals(target, referenceTarget) ||
                !ReferenceEquals(target.Parent, body) || !IsReferenceOnlyRun(referenceRun) ||
                !CoversWholeParagraph(target, matchingStarts[0], matchingEnds[0], referenceRun))
            {
                reason = $"comment {nativeId} is not an exact whole top-level paragraph anchor";
                return false;
            }
            var bodyIndex = checked((uint)body.ChildElements.ToList().IndexOf(target));
            if (!blockByBodyIndex.TryGetValue(bodyIndex, out var block))
            {
                reason = $"comment {nativeId} target does not resolve to a modeled body block";
                return false;
            }
            rootAnchors.Add(nativeId, (target, matchingStarts[0], matchingEnds[0], referenceRun, bodyIndex, block));
        }

        var sources = new List<DocxClassicCommentSource>(elements.Length);
        for (var index = 0; index < elements.Length; index++)
        {
            var element = elements[index];
            var nativeId = element.Id!.Value!;
            var extendedInfo = extended.IsModern ? extended.ByNativeCommentId[nativeId] : null;
            var rootNativeId = nativeId;
            var parentModelId = string.Empty;
            if (extendedInfo?.ParentParagraphId.Length > 0)
            {
                if (!nativeIdByParagraphId.TryGetValue(extendedInfo.ParentParagraphId, out rootNativeId) || !modelIdByNativeId.TryGetValue(rootNativeId, out parentModelId))
                {
                    reason = $"comment {nativeId} parent does not resolve to a root comment";
                    return false;
                }
            }
            if (!rootAnchors.TryGetValue(rootNativeId, out var anchor))
            {
                reason = $"comment {nativeId} does not resolve to a root anchor";
                return false;
            }
            var artifact = ReadArtifact(
                part.Comments,
                element,
                $"document/comment/{index + 1}",
                anchor.Block.Id,
                anchor.BodyIndex,
                anchor.Start,
                anchor.End,
                anchor.ReferenceRun,
                extended,
                extendedInfo,
                parentModelId);
            sources.Add(new DocxClassicCommentSource(
                element,
                anchor.Target,
                anchor.Start,
                anchor.End,
                anchor.ReferenceRun,
                artifact));
        }

        graph = new DocxClassicCommentGraph(part, sources, extended);
        return true;
    }

    private static DocumentComment ReadArtifact(
        W.Comments commentsRoot,
        W.Comment element,
        string id,
        string targetBlockId,
        uint bodyIndex,
        W.CommentRangeStart start,
        W.CommentRangeEnd end,
        W.Run referenceRun,
        DocxExtendedCommentGraph extended,
        DocxExtendedCommentInfo? extendedInfo,
        string parentCommentId)
    {
        var artifact = new DocumentComment
        {
            Id = id,
            TargetBlockId = targetBlockId,
            Author = AttributeValue(element, "author") ?? string.Empty,
            Text = element.Descendants<W.Text>().Single().Text,
        };
        var initials = AttributeValue(element, "initials");
        if (initials is not null) artifact.Initials = initials;
        var createdAt = AttributeValue(element, "date");
        if (createdAt is not null) artifact.CreatedAt = createdAt;
        if (extendedInfo is not null)
        {
            artifact.ParentCommentId = parentCommentId;
            artifact.Resolved = extendedInfo.Resolved;
            artifact.ParagraphId = extendedInfo.ParagraphId;
            artifact.DurableId = extendedInfo.DurableId;
            if (extendedInfo.DateUtc.Length > 0) artifact.DateUtc = extendedInfo.DateUtc;
            if (extendedInfo.Person is not null) artifact.Person = extendedInfo.Person.Clone();
            if (extendedInfo.IntelligentPlaceholder) artifact.IntelligentPlaceholder = true;
        }
        artifact.Source = new DocumentCommentSourceBinding
        {
            NativeCommentId = element.Id?.Value ?? string.Empty,
            TargetBodyIndex = bodyIndex,
            CommentElementSha256 = HashElement(element),
            ResidualSha256 = GraphResidualHash(commentsRoot),
            AnchorSha256 = AnchorHash(start, end, referenceRun),
            Editable = true,
            ExtendedGraphSha256 = extended.ExtendedGraphSha256,
            CommentsIdsGraphSha256 = extended.CommentsIdsGraphSha256,
            CommentsExtensibleGraphSha256 = extended.CommentsExtensibleGraphSha256,
            PeopleGraphSha256 = extended.PeopleGraphSha256,
            ThreadEditable = extended.IsModern,
        };
        artifact.Source.SemanticSha256 = SemanticHash(artifact);
        return artifact;
    }

    private static W.Comment BuildComment(DocumentComment artifact, string nativeId)
    {
        var comment = new W.Comment { Id = nativeId };
        ApplyValues(comment, artifact);
        return comment;
    }

    private static void ApplyValues(W.Comment comment, DocumentComment artifact)
    {
        SetAttribute(comment, "author", artifact.Author);
        SetOptionalAttribute(comment, "initials", artifact.HasInitials ? artifact.Initials : null);
        SetOptionalAttribute(comment, "date", artifact.HasCreatedAt ? artifact.CreatedAt : null);
        var text = comment.Descendants<W.Text>().SingleOrDefault();
        if (text is null)
        {
            comment.RemoveAllChildren();
            text = Text(artifact.Text);
            comment.Append(new W.Paragraph(new W.Run(text)));
        }
        else
        {
            text.Text = artifact.Text;
            text.Space = artifact.Text.Length != artifact.Text.Trim().Length ? SpaceProcessingModeValues.Preserve : null;
        }
    }

    private static void AddWholeParagraphAnchors(
        W.Paragraph paragraph,
        IReadOnlyList<(DocumentComment Artifact, string NativeId, W.Paragraph Target)> comments)
    {
        var firstContent = paragraph.ChildElements.FirstOrDefault(child => child is not W.ParagraphProperties);
        OpenXmlElement? cursor = paragraph.ChildElements.LastOrDefault(child => child is not W.ParagraphProperties);
        foreach (var item in comments)
        {
            var start = new W.CommentRangeStart { Id = item.NativeId };
            if (firstContent is null) paragraph.Append(start);
            else paragraph.InsertBefore(start, firstContent);
            cursor ??= start;
        }
        foreach (var item in comments.Reverse())
        {
            var end = new W.CommentRangeEnd { Id = item.NativeId };
            paragraph.InsertAfter(end, cursor);
            cursor = end;
        }
        foreach (var item in comments)
        {
            var reference = new W.Run(new W.CommentReference { Id = item.NativeId });
            paragraph.InsertAfter(reference, cursor);
            cursor = reference;
        }
    }

    private static bool IsSimpleCommentBody(W.Comment comment)
    {
        if (comment.ChildElements.Count != 1 || comment.FirstChild is not W.Paragraph paragraph) return false;
        if (paragraph.ChildElements.Any(child => child is not W.ParagraphProperties and not W.Run)) return false;
        if (paragraph.Elements<W.ParagraphProperties>().Count() > 1) return false;
        var runs = paragraph.Elements<W.Run>().ToArray();
        if (runs.Length != 1 || runs[0].ChildElements.Any(child => child is not W.RunProperties and not W.Text)) return false;
        return runs[0].Elements<W.RunProperties>().Count() <= 1 && runs[0].Elements<W.Text>().Count() == 1;
    }

    private static bool IsReferenceOnlyRun(W.Run run) =>
        run.ChildElements.All(child => child is W.RunProperties or W.CommentReference) &&
        run.Elements<W.RunProperties>().Count() <= 1 &&
        run.Elements<W.CommentReference>().Count() == 1;

    private static bool CoversWholeParagraph(
        W.Paragraph paragraph,
        W.CommentRangeStart start,
        W.CommentRangeEnd end,
        W.Run referenceRun)
    {
        var children = paragraph.ChildElements.ToArray();
        var startIndex = Array.IndexOf(children, start);
        var endIndex = Array.IndexOf(children, end);
        var referenceIndex = Array.IndexOf(children, referenceRun);
        if (startIndex < 0 || endIndex <= startIndex || referenceIndex <= endIndex) return false;
        var contentIndexes = children
            .Select((child, index) => (child, index))
            .Where(item => item.child is not W.ParagraphProperties &&
                           item.child is not W.CommentRangeStart &&
                           item.child is not W.CommentRangeEnd &&
                           (item.child is not W.Run run || !IsReferenceOnlyRun(run)))
            .Select(item => item.index)
            .ToArray();
        return contentIndexes.Length == 0 ||
               (startIndex < contentIndexes.Min() && endIndex > contentIndexes.Max());
    }

    private static bool ClassicValuesEqual(DocumentComment left, DocumentComment right) =>
        left.Author == right.Author && left.Text == right.Text &&
        left.HasInitials == right.HasInitials && left.Initials == right.Initials &&
        left.HasCreatedAt == right.HasCreatedAt && left.CreatedAt == right.CreatedAt;

    private static string SemanticHash(DocumentComment comment)
    {
        var semantic = comment.Clone();
        semantic.Id = string.Empty;
        semantic.Source = null;
        return Hash(semantic.ToByteArray());
    }

    private static string GraphResidualHash(W.Comments comments)
    {
        var residual = (W.Comments)comments.CloneNode(true);
        foreach (var comment in residual.Elements<W.Comment>())
        {
            comment.RemoveAttribute("author", WordprocessingNamespace);
            comment.RemoveAttribute("initials", WordprocessingNamespace);
            comment.RemoveAttribute("date", WordprocessingNamespace);
            var text = comment.Descendants<W.Text>().Single();
            text.Text = string.Empty;
            text.Space = null;
        }
        return HashElement(residual);
    }

    private static string AnchorHash(W.CommentRangeStart start, W.CommentRangeEnd end, W.Run referenceRun) =>
        Hash(Encoding.UTF8.GetBytes($"{start.OuterXml}\0{end.OuterXml}\0{referenceRun.OuterXml}"));

    private static string HashElement(OpenXmlElement element) => Hash(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static string? AttributeValue(OpenXmlElement element, string localName)
    {
        var attribute = element.GetAttributes().FirstOrDefault(item =>
            item.LocalName == localName && item.NamespaceUri == WordprocessingNamespace);
        return string.IsNullOrEmpty(attribute.LocalName) ? null : attribute.Value;
    }

    private static void SetAttribute(OpenXmlElement element, string localName, string value) =>
        element.SetAttribute(new OpenXmlAttribute("w", localName, WordprocessingNamespace, value));

    private static void SetOptionalAttribute(OpenXmlElement element, string localName, string? value)
    {
        if (value is null) element.RemoveAttribute(localName, WordprocessingNamespace);
        else SetAttribute(element, localName, value);
    }

    private static W.Text Text(string value) => new(value)
    {
        Space = value.Length != value.Trim().Length ? SpaceProcessingModeValues.Preserve : null,
    };

    private static void ValidateText(string value, string name, int minimumLength, int maximumLength)
    {
        if (value.Length < minimumLength || value.Length > maximumLength)
            throw new CodecException(
                "invalid_document_comment",
                $"{name} must contain {minimumLength} through {maximumLength} characters.");
        try
        {
            XmlConvert.VerifyXmlChars(value);
        }
        catch (XmlException exception)
        {
            throw new CodecException(
                "invalid_document_comment",
                $"{name} contains characters that cannot be represented in XML.",
                innerException: exception);
        }
    }

    private static void ValidateOptionalHex(string value, string name)
    {
        if (value.Length == 0) return;
        if (value.Length != 8 || value.Any(character => !Uri.IsHexDigit(character)))
            throw new CodecException("invalid_document_comment", $"{name} must contain exactly eight hexadecimal digits when present.");
    }

    private static void ValidateOptionalDurableId(string value, string name)
    {
        if (value.Length == 0) return;
        ValidateOptionalHex(value, name);
        if (!uint.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var number) ||
            number == 0 || number >= 0x7FFFFFFF)
            throw new CodecException("invalid_document_comment", $"{name} must be between 00000001 and 7FFFFFFE when present.");
    }

    private static void ValidateOptionalParagraphId(string value, string name)
    {
        if (value.Length == 0) return;
        ValidateOptionalHex(value, name);
        if (!uint.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var number) ||
            number == 0 || number >= 0x80000000)
            throw new CodecException("invalid_document_comment", $"{name} must be between 00000001 and 7FFFFFFF when present.");
    }
}
