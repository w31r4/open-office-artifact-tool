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
    IReadOnlyList<DocxClassicCommentSource> Comments);

// Owns a deliberately bounded classic WordprocessingML comment profile:
// one top-level body paragraph is anchored by exactly one start/end/reference
// triplet, and the comment body contains one paragraph/run/text. Rich or
// extended comment graphs remain in the opaque source package and fail closed.
internal static class DocxClassicCommentCodec
{
    private const string WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const int MaxAuthorLength = 255;
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
                $"Preserved a DOCX comment graph that is outside the editable classic-comment profile: {reason}",
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
        foreach (var comment in document.Comments)
        {
            if (string.IsNullOrWhiteSpace(comment.Id) || !ids.Add(comment.Id))
                throw new CodecException("invalid_document_comment", "Document comments require unique, non-empty IDs.");
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
            }
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

        var authored = new List<(DocumentComment Artifact, string NativeId, W.Paragraph Target)>();
        for (var index = 0; index < document.Comments.Count; index++)
        {
            var artifact = document.Comments[index];
            if (!targets.TryGetValue(artifact.TargetBlockId, out var target))
                throw new CodecException(
                    "unsupported_document_comment_target",
                    $"Document comment {artifact.Id} does not resolve to an authored top-level paragraph.");
            var nativeId = index.ToString(CultureInfo.InvariantCulture);
            part.Comments.Append(BuildComment(artifact, nativeId));
            authored.Add((artifact, nativeId, target));
        }

        foreach (var group in authored.GroupBy(item => item.Target))
            AddWholeParagraphAnchors(group.Key, group.ToArray());
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
        var changed = false;
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
            AssertBindingMatches(requested, binding, source, original);

            var requestedHash = SemanticHash(requested);
            if (requestedHash.Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase)) continue;
            if (!binding.Editable || requested.Id != original.Id || requested.TargetBlockId != original.TargetBlockId)
                throw new CodecException(
                    "unsupported_document_comment_edit",
                    $"Document comment {requested.Id} changes source-bound identity or target topology.",
                    "word/comments.xml");

            ApplyValues(source.Element, requested);
            if (!ResidualHash(source.Element).Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "document_comment_residual_not_preserved",
                    $"Document comment {requested.Id} changed unmodeled comment formatting or extension markup.",
                    "word/comments.xml");
            if (!AnchorHash(source.Start, source.End, source.ReferenceRun).Equals(binding.AnchorSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "document_comment_anchor_not_preserved",
                    $"Document comment {requested.Id} changed its source anchor triplet.",
                    "word/document.xml");
            var verified = ReadArtifact(
                source.Element,
                requested.Id,
                original.TargetBlockId,
                original.Source.TargetBodyIndex,
                source.Start,
                source.End,
                source.ReferenceRun);
            if (!SemanticHash(verified).Equals(requestedHash, StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "document_comment_semantics_not_applied",
                    $"Document comment {requested.Id} does not match the requested semantics after editing.",
                    "word/comments.xml");
            changed = true;
        }

        if (!changed) return;
        graph.Part.Comments!.Save();
        context.MarkCommentsMutated(graph.Part);
    }

    private static void AssertBindingMatches(
        DocumentComment requested,
        DocumentCommentSourceBinding binding,
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
            !HashElement(source.Element).Equals(binding.CommentElementSha256, StringComparison.OrdinalIgnoreCase) ||
            !ResidualHash(source.Element).Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase) ||
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
        if (context.Owner.Parts.Any(pair => IsExtendedCommentRelationship(pair.OpenXmlPart.RelationshipType)))
        {
            reason = "extended, durable, or people comment relationships are present";
            return false;
        }
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
        foreach (var element in elements)
        {
            var nativeId = element.Id?.Value;
            if (string.IsNullOrWhiteSpace(nativeId) ||
                !int.TryParse(nativeId, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var numericId) ||
                numericId < 0 || !nativeIds.Add(nativeId))
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

        var starts = body.Descendants<W.CommentRangeStart>().ToArray();
        var ends = body.Descendants<W.CommentRangeEnd>().ToArray();
        var references = body.Descendants<W.CommentReference>().ToArray();
        if (starts.Length != elements.Length || ends.Length != elements.Length || references.Length != elements.Length)
        {
            reason = "comment/start/end/reference counts do not match";
            return false;
        }
        if (starts.Any(item => !nativeIds.Contains(item.Id?.Value ?? string.Empty)) ||
            ends.Any(item => !nativeIds.Contains(item.Id?.Value ?? string.Empty)) ||
            references.Any(item => !nativeIds.Contains(item.Id?.Value ?? string.Empty)))
        {
            reason = "the body contains a dangling or foreign comment anchor";
            return false;
        }

        var blockByBodyIndex = document.Blocks
            .Where(block => block.Source is not null)
            .ToDictionary(block => block.Source.BodyIndex);
        var sources = new List<DocxClassicCommentSource>(elements.Length);
        for (var index = 0; index < elements.Length; index++)
        {
            var element = elements[index];
            var nativeId = element.Id!.Value!;
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
            var artifact = ReadArtifact(
                element,
                $"document/comment/{index + 1}",
                block.Id,
                bodyIndex,
                matchingStarts[0],
                matchingEnds[0],
                referenceRun);
            sources.Add(new DocxClassicCommentSource(
                element,
                target,
                matchingStarts[0],
                matchingEnds[0],
                referenceRun,
                artifact));
        }

        graph = new DocxClassicCommentGraph(part, sources);
        return true;
    }

    private static DocumentComment ReadArtifact(
        W.Comment element,
        string id,
        string targetBlockId,
        uint bodyIndex,
        W.CommentRangeStart start,
        W.CommentRangeEnd end,
        W.Run referenceRun)
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
        artifact.Source = new DocumentCommentSourceBinding
        {
            NativeCommentId = element.Id?.Value ?? string.Empty,
            TargetBodyIndex = bodyIndex,
            CommentElementSha256 = HashElement(element),
            ResidualSha256 = ResidualHash(element),
            AnchorSha256 = AnchorHash(start, end, referenceRun),
            Editable = true,
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

    private static bool IsExtendedCommentRelationship(string relationshipType) =>
        relationshipType.EndsWith("/commentsExtended", StringComparison.OrdinalIgnoreCase) ||
        relationshipType.EndsWith("/commentsIds", StringComparison.OrdinalIgnoreCase) ||
        relationshipType.EndsWith("/commentsExtensible", StringComparison.OrdinalIgnoreCase) ||
        relationshipType.EndsWith("/people", StringComparison.OrdinalIgnoreCase);

    private static string SemanticHash(DocumentComment comment)
    {
        var semantic = comment.Clone();
        semantic.Id = string.Empty;
        semantic.Source = null;
        return Hash(semantic.ToByteArray());
    }

    private static string ResidualHash(W.Comment comment)
    {
        var residual = (W.Comment)comment.CloneNode(true);
        residual.RemoveAttribute("author", WordprocessingNamespace);
        residual.RemoveAttribute("initials", WordprocessingNamespace);
        residual.RemoveAttribute("date", WordprocessingNamespace);
        var text = residual.Descendants<W.Text>().Single();
        text.Text = string.Empty;
        text.Space = null;
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
}
