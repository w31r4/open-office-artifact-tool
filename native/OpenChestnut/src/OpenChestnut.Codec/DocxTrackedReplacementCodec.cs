using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

internal sealed record DocxTrackedReplacementOutput(
    byte[] File,
    DocumentTrackedReplacementResult Result,
    IReadOnlyList<Diagnostic> Diagnostics);

internal sealed record DocxInlineTrackedReplacement(
    W.DeletedRun Deletion,
    W.InsertedRun Insertion,
    string DeletedText,
    string InsertedText);

// Adds one exact, direct-run replacement to an existing DOCX. This operation
// deliberately owns package mutation instead of expanding the general
// DocumentParagraph model with source-only revision topology.
internal static class DocxTrackedReplacementCodec
{
    private const string DocumentPath = "word/document.xml";

    internal static DocxTrackedReplacementOutput Add(
        byte[] sourceBytes,
        DocumentTrackedReplacementRequest request,
        EffectiveCodecLimits limits)
    {
        ValidateRequest(request);
        var sourceHash = Hash(sourceBytes);
        if (!IsSha256(request.ExpectedSourceSha256) ||
            !sourceHash.Equals(request.ExpectedSourceSha256, StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "document_source_hash_mismatch",
                "DOCX tracked replacement requires expected_source_sha256 to match the exact input bytes.");

        var sourceOpaque = PackageGuards.ValidateAndCollectOpaque(sourceBytes, limits, OpcPackageProfile.Docx, includeSourcePackage: false);
        var sourceInventory = DocxRevisionMarkup.Inventory(sourceBytes);
        var revisionCount = sourceInventory.Values.Aggregate(0UL, (sum, value) => checked(sum + (ulong)value));
        if (revisionCount > limits.MaxCells)
            throw new CodecException(
                "document_item_budget_exceeded",
                $"DOCX contains {revisionCount} revision markers and exceeds max_cells ({limits.MaxCells}).");
        var (deletionId, insertionId) = DocxRevisionMarkup.AllocatePair(sourceBytes, limits.MaxCells);

        using var stream = new MemoryStream();
        stream.Write(sourceBytes);
        stream.Position = 0;
        uint bodyIndex;
        string sourceElementHash;
        using (var package = WordprocessingDocument.Open(stream, isEditable: true, new OpenSettings { AutoSave = false }))
        {
            var mainPart = package.MainDocumentPart ??
                throw new CodecException("missing_document_part", "DOCX package has no Main Document part.", DocumentPath);
            var body = mainPart.Document?.Body ??
                throw new CodecException("missing_document_body", "DOCX package has no document body.", DocumentPath);
            var blocks = body.ChildElements.Where(element => element is not W.SectionProperties).ToArray();
            if (request.TargetBlockIndex >= blocks.Length)
                throw new CodecException(
                    "document_tracked_replacement_target_not_found",
                    $"DOCX tracked replacement target block {request.TargetBlockIndex} is outside the {blocks.Length}-block document body.",
                    DocumentPath);
            if (blocks[request.TargetBlockIndex] is not W.Paragraph paragraph)
                throw new CodecException(
                    "unsupported_document_tracked_replacement_target",
                    $"DOCX tracked replacement target block {request.TargetBlockIndex} is not a paragraph.",
                    DocumentPath);

            bodyIndex = BodyIndex(body, paragraph);
            sourceElementHash = HashElement(paragraph);
            var target = FindTarget(paragraph, request.ExpectedParagraphText, request.Search);
            var date = Date(request);
            var prefix = target.Text.Text[..target.MatchIndex];
            var suffix = target.Text.Text[(target.MatchIndex + request.Search.Length)..];
            if (prefix.Length > 0) target.Run.InsertBeforeSelf(CloneRun(target.Run, prefix, deleted: false));
            var deletion = new W.DeletedRun(CloneRun(target.Run, request.Search, deleted: true))
            {
                Id = deletionId,
                Author = request.Author,
                Date = date,
            };
            target.Run.InsertBeforeSelf(deletion);
            var insertion = new W.InsertedRun(CloneRun(target.Run, request.Replacement, deleted: false))
            {
                Id = insertionId,
                Author = request.Author,
                Date = date,
            };
            target.Run.InsertBeforeSelf(insertion);
            if (suffix.Length > 0) target.Run.InsertBeforeSelf(CloneRun(target.Run, suffix, deleted: false));
            target.Run.Remove();
            mainPart.Document!.Save();
        }

        var outputBytes = stream.ToArray();
        DocxCodec.ValidateOutputBudget(outputBytes, limits);
        var retainedValidationErrorCount = DocxCodec.ValidateOffice2021AgainstSource(sourceBytes, outputBytes);
        var outputInventory = DocxRevisionMarkup.Inventory(outputBytes);
        foreach (var (path, count) in sourceInventory)
        {
            var expected = count + (path.Equals(DocumentPath, StringComparison.OrdinalIgnoreCase) ? 2 : 0);
            if (outputInventory.GetValueOrDefault(path) != expected)
                throw new CodecException(
                    "document_tracked_replacement_scope_violation",
                    $"DOCX tracked replacement changed the revision inventory for {path} unexpectedly.",
                    path);
        }
        if (outputInventory.Keys.Any(path => !sourceInventory.ContainsKey(path)))
            throw new CodecException(
                "document_tracked_replacement_scope_violation",
                "DOCX tracked replacement introduced an unexpected WordprocessingML part.");

        var changedParts = ChangedParts(sourceBytes, outputBytes);
        if (changedParts.Length != 1 || !changedParts[0].Equals(DocumentPath, StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "document_tracked_replacement_scope_violation",
                $"DOCX tracked replacement changed unexpected OPC parts: {string.Join(", ", changedParts)}.");

        var outputOpaque = PackageGuards.ValidateAndCollectOpaque(outputBytes, limits, OpcPackageProfile.Docx, includeSourcePackage: false);
        PackageGuards.AssertOpaqueGraphMatches(sourceOpaque, outputOpaque, "opaque_content_not_preserved");

        string outputElementHash;
        using (var outputStream = new MemoryStream(outputBytes, writable: false))
        using (var package = WordprocessingDocument.Open(outputStream, isEditable: false))
        {
            var body = package.MainDocumentPart?.Document?.Body ??
                throw new CodecException("missing_document_body", "Rewritten DOCX package has no document body.", DocumentPath);
            var blocks = body.ChildElements.Where(element => element is not W.SectionProperties).ToArray();
            if (request.TargetBlockIndex >= blocks.Length ||
                blocks[request.TargetBlockIndex] is not W.Paragraph paragraph ||
                !TryReadFinalizable(paragraph, out var replacement) ||
                replacement.Deletion.Id?.Value != deletionId ||
                replacement.Insertion.Id?.Value != insertionId ||
                replacement.DeletedText != request.Search ||
                replacement.InsertedText != request.Replacement ||
                AcceptedText(paragraph) != request.ExpectedParagraphText.Replace(request.Search, request.Replacement, StringComparison.Ordinal))
                throw new CodecException(
                    "document_tracked_replacement_verification_failed",
                    "Rewritten DOCX does not contain the requested bounded tracked replacement.",
                    DocumentPath);
            outputElementHash = HashElement(paragraph);
        }

        var result = new DocumentTrackedReplacementResult
        {
            SourceSha256 = sourceHash,
            OutputSha256 = Hash(outputBytes),
            TargetBlockIndex = request.TargetBlockIndex,
            TargetBodyIndex = bodyIndex,
            SourceElementSha256 = sourceElementHash,
            OutputElementSha256 = outputElementHash,
            DeletedTextSha256 = Hash(Encoding.UTF8.GetBytes(request.Search)),
            InsertedTextSha256 = Hash(Encoding.UTF8.GetBytes(request.Replacement)),
            DeletedTextChars = checked((uint)request.Search.Length),
            InsertedTextChars = checked((uint)request.Replacement.Length),
            DeletionNativeRevisionId = deletionId,
            InsertionNativeRevisionId = insertionId,
        };
        result.ChangedParts.Add(changedParts);
        var diagnostics = new List<Diagnostic>();
        if (retainedValidationErrorCount > 0)
            diagnostics.Add(CodecProtocol.Warning(
                "source_openxml_validation_warnings_preserved",
                $"Preserved {retainedValidationErrorCount} pre-existing Office 2021 validation warning(s) from the source package; tracked replacement introduced none."));
        return new DocxTrackedReplacementOutput(outputBytes, result, diagnostics);
    }

    internal static bool TryReadFinalizable(W.Paragraph paragraph, out DocxInlineTrackedReplacement replacement)
    {
        replacement = null!;
        var children = paragraph.ChildElements.Where(child => child is not W.ParagraphProperties).ToArray();
        if (children.Any(child => child is not W.Run and not W.DeletedRun and not W.InsertedRun)) return false;
        var deletionIndexes = children.Select((child, index) => (child, index)).Where(item => item.child is W.DeletedRun).ToArray();
        var insertionIndexes = children.Select((child, index) => (child, index)).Where(item => item.child is W.InsertedRun).ToArray();
        if (deletionIndexes.Length != 1 || insertionIndexes.Length != 1 || insertionIndexes[0].index != deletionIndexes[0].index + 1)
            return false;
        if (children.Where(child => child is W.Run).Cast<W.Run>().Any(run => !TryReadTextRun(run, deleted: false, out _)))
            return false;

        var deletion = (W.DeletedRun)deletionIndexes[0].child;
        var insertion = (W.InsertedRun)insertionIndexes[0].child;
        if (!TryReadRevisionRun(deletion, deleted: true, out var deletedText) ||
            !TryReadRevisionRun(insertion, deleted: false, out var insertedText) ||
            deletedText.Length == 0 || insertedText.Length == 0 ||
            string.IsNullOrWhiteSpace(deletion.Id?.Value) || string.IsNullOrWhiteSpace(insertion.Id?.Value) ||
            deletion.Id?.Value == insertion.Id?.Value ||
            string.IsNullOrWhiteSpace(deletion.Author?.Value) || string.IsNullOrWhiteSpace(insertion.Author?.Value) ||
            deletion.Author?.Value != insertion.Author?.Value || deletion.Date?.Value != insertion.Date?.Value)
            return false;

        replacement = new DocxInlineTrackedReplacement(deletion, insertion, deletedText, insertedText);
        return true;
    }

    private sealed record TargetRun(W.Run Run, W.Text Text, int MatchIndex);

    private static TargetRun FindTarget(W.Paragraph paragraph, string expectedText, string search)
    {
        var children = paragraph.ChildElements.Where(child => child is not W.ParagraphProperties).ToArray();
        if (children.Length == 0 || children.Any(child => child is not W.Run))
            throw Unsupported("Target paragraph must contain only direct ordinary text runs and no fields, hyperlinks, content controls, drawings, or existing revisions.");

        var runs = children.Cast<W.Run>().ToArray();
        var textNodes = new List<(W.Run Run, W.Text Text)>();
        foreach (var run in runs)
        {
            if (!TryReadTextRun(run, deleted: false, out _) || run.Elements<W.Text>().Count() != 1)
                throw Unsupported("Target paragraph must contain exactly one plain w:t node in each direct run.");
            textNodes.Add((run, run.Elements<W.Text>().Single()));
        }
        var paragraphText = string.Concat(textNodes.Select(item => item.Text.Text));
        if (!paragraphText.Equals(expectedText, StringComparison.Ordinal))
            throw new CodecException(
                "document_tracked_replacement_text_mismatch",
                "DOCX tracked replacement expected_paragraph_text does not match the exact target paragraph.",
                DocumentPath);

        var matches = new List<TargetRun>();
        foreach (var item in textNodes)
        {
            for (var index = item.Text.Text.IndexOf(search, StringComparison.Ordinal);
                 index >= 0;
                 index = item.Text.Text.IndexOf(search, index + 1, StringComparison.Ordinal))
                matches.Add(new TargetRun(item.Run, item.Text, index));
        }
        if (matches.Count == 0 && paragraphText.Contains(search, StringComparison.Ordinal))
            throw new CodecException(
                "document_tracked_replacement_cross_run_match",
                "DOCX tracked replacement search text crosses native run boundaries; this bounded operation requires one direct w:t match.",
                DocumentPath);
        if (matches.Count != 1)
            throw new CodecException(
                "document_tracked_replacement_match_not_unique",
                $"DOCX tracked replacement requires exactly one direct w:t match; found {matches.Count}.",
                DocumentPath);
        return matches[0];
    }

    private static W.Run CloneRun(W.Run source, string value, bool deleted)
    {
        var run = (W.Run)source.CloneNode(true);
        var text = run.Elements<W.Text>().Single();
        OpenXmlElement replacement = deleted
            ? new W.DeletedText(value) { Space = Preserve(value) }
            : new W.Text(value) { Space = Preserve(value) };
        text.InsertAfterSelf(replacement);
        text.Remove();
        return run;
    }

    private static bool TryReadRevisionRun(OpenXmlCompositeElement revision, bool deleted, out string text)
    {
        text = string.Empty;
        if (revision.ChildElements.Count != 1 || revision.FirstChild is not W.Run run) return false;
        return TryReadTextRun(run, deleted, out text);
    }

    private static bool TryReadTextRun(W.Run run, bool deleted, out string text)
    {
        text = string.Empty;
        if (run.ChildElements.Any(child => child is not W.RunProperties &&
                (deleted ? child is not W.DeletedText : child is not W.Text)))
            return false;
        if (deleted)
        {
            var values = run.Elements<W.DeletedText>().ToArray();
            if (values.Length != 1) return false;
            text = values[0].Text;
        }
        else
        {
            var values = run.Elements<W.Text>().ToArray();
            if (values.Length != 1) return false;
            text = values[0].Text;
        }
        return true;
    }

    private static string AcceptedText(W.Paragraph paragraph) =>
        string.Concat(paragraph.ChildElements.Where(child => child is not W.DeletedRun)
            .SelectMany(child => child.Descendants<W.Text>())
            .Select(value => value.Text));

    private static void ValidateRequest(DocumentTrackedReplacementRequest? request)
    {
        if (request is null)
            throw new CodecException("missing_tracked_replacement", "DOCX tracked replacement requires tracked_replacement options.");
        ValidateText(request.ExpectedParagraphText, "expected paragraph text");
        ValidateText(request.Search, "search text");
        ValidateText(request.Replacement, "replacement text");
        if (string.IsNullOrWhiteSpace(request.Author) || request.Author.Length > 255 || request.Author.Any(char.IsControl))
            throw new CodecException("invalid_document_tracked_replacement", "DOCX tracked replacement author must contain 1 through 255 characters without controls.");
        if (request.HasDate && (!DateTimeOffset.TryParse(request.Date, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out _) || request.Date.Length > 64))
            throw new CodecException("invalid_document_tracked_replacement", "DOCX tracked replacement date must be an ISO 8601 timestamp of at most 64 characters.");
    }

    private static void ValidateText(string value, string label)
    {
        if (value.Length == 0 || value.Length > 1_000_000)
            throw new CodecException("invalid_document_tracked_replacement", $"DOCX tracked replacement {label} must contain 1 through 1,000,000 characters.");
        try
        {
            System.Xml.XmlConvert.VerifyXmlChars(value);
        }
        catch (System.Xml.XmlException exception)
        {
            throw new CodecException("invalid_document_tracked_replacement", $"DOCX tracked replacement {label} must contain only XML-safe characters.", innerException: exception);
        }
    }

    private static DateTime? Date(DocumentTrackedReplacementRequest request) => request.HasDate
        ? DateTimeOffset.Parse(request.Date, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind).UtcDateTime
        : null;

    private static SpaceProcessingModeValues? Preserve(string value) =>
        value.Length != value.Trim().Length ? SpaceProcessingModeValues.Preserve : null;

    private static uint BodyIndex(W.Body body, W.Paragraph paragraph)
    {
        for (var index = 0; index < body.ChildElements.Count; index++)
            if (ReferenceEquals(body.ChildElements[index], paragraph)) return checked((uint)index);
        throw new CodecException("document_tracked_replacement_target_not_found", "DOCX tracked replacement target paragraph is not in the document body.", DocumentPath);
    }

    private static string[] ChangedParts(byte[] sourceBytes, byte[] outputBytes)
    {
        var source = PartHashes(sourceBytes);
        var output = PartHashes(outputBytes);
        if (!source.Keys.OrderBy(item => item, StringComparer.OrdinalIgnoreCase)
                .SequenceEqual(output.Keys.OrderBy(item => item, StringComparer.OrdinalIgnoreCase), StringComparer.OrdinalIgnoreCase))
            throw new CodecException("document_tracked_replacement_scope_violation", "DOCX tracked replacement changed the OPC part inventory.");
        return source.Keys
            .Where(path => !source[path].Equals(output[path], StringComparison.Ordinal))
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static Dictionary<string, string> PartHashes(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in archive.Entries.Where(entry => !entry.FullName.EndsWith("/", StringComparison.Ordinal)))
        {
            using var partStream = entry.Open();
            using var copy = new MemoryStream();
            partStream.CopyTo(copy);
            result.Add(entry.FullName, Hash(copy.ToArray()));
        }
        return result;
    }

    private static string HashElement(OpenXmlElement element) => Hash(Encoding.UTF8.GetBytes(element.OuterXml));
    private static bool IsSha256(string value) => value.Length == 64 && value.All(Uri.IsHexDigit);
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static CodecException Unsupported(string message) => new("unsupported_document_tracked_replacement_topology", message, DocumentPath);
}
