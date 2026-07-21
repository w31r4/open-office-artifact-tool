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
    string InsertedText,
    uint DeletedRunCount,
    uint InsertedRunCount);

internal sealed record DocxResolvedTrackedReplacementTarget(
    W.Paragraph Paragraph,
    uint BodyIndex,
    DocumentTrackedReplacementTarget Selector);

// Adds one exact direct-text replacement to an existing DOCX. A literal may
// cross adjacent ordinary runs only when their run properties are identical.
// This operation deliberately owns package mutation instead of expanding the
// general DocumentParagraph model with source-only revision topology.
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
        uint matchedSourceRunCount;
        DocumentTrackedReplacementTarget selector;
        string sourceElementHash;
        using (var package = WordprocessingDocument.Open(stream, isEditable: true, new OpenSettings { AutoSave = false }))
        {
            var mainPart = package.MainDocumentPart ??
                throw new CodecException("missing_document_part", "DOCX package has no Main Document part.", DocumentPath);
            var body = mainPart.Document?.Body ??
                throw new CodecException("missing_document_body", "DOCX package has no document body.", DocumentPath);
            var resolved = ResolveTarget(body, request);
            var paragraph = resolved.Paragraph;
            bodyIndex = resolved.BodyIndex;
            selector = resolved.Selector;
            sourceElementHash = HashElement(paragraph);
            var target = FindTarget(paragraph, request.ExpectedParagraphText, request.Search);
            matchedSourceRunCount = checked((uint)target.Segments.Count);
            var date = Date(request);
            var first = target.Segments[0];
            var last = target.Segments[^1];
            var prefix = first.Text.Text[..first.Start];
            var suffix = last.Text.Text[(last.Start + last.Length)..];
            if (prefix.Length > 0) first.Run.InsertBeforeSelf(CloneRun(first.Run, prefix, deleted: false));
            var deletion = new W.DeletedRun
            {
                Id = deletionId,
                Author = request.Author,
                Date = date,
            };
            foreach (var segment in target.Segments)
                deletion.Append(CloneRun(
                    segment.Run,
                    segment.Text.Text.Substring(segment.Start, segment.Length),
                    deleted: true));
            first.Run.InsertBeforeSelf(deletion);
            var insertion = new W.InsertedRun(CloneRun(first.Run, request.Replacement, deleted: false))
            {
                Id = insertionId,
                Author = request.Author,
                Date = date,
            };
            first.Run.InsertBeforeSelf(insertion);
            if (suffix.Length > 0) first.Run.InsertBeforeSelf(CloneRun(last.Run, suffix, deleted: false));
            foreach (var segment in target.Segments) segment.Run.Remove();
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
            var resolved = ResolveTarget(body, request);
            var paragraph = resolved.Paragraph;
            if (!SameTarget(resolved.Selector, selector) ||
                !TryReadFinalizable(paragraph, out var replacement) ||
                replacement.Deletion.Id?.Value != deletionId ||
                replacement.Insertion.Id?.Value != insertionId ||
                replacement.DeletedText != request.Search ||
                replacement.InsertedText != request.Replacement ||
                replacement.DeletedRunCount != matchedSourceRunCount ||
                replacement.InsertedRunCount != 1 ||
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
            Target = selector.Clone(),
            MatchedSourceRunCount = matchedSourceRunCount,
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
        if (!TryReadRevisionRuns(deletion, deleted: true, out var deletedText, out var deletionRuns) ||
            !TryReadRevisionRuns(insertion, deleted: false, out var insertedText, out var insertionRuns) ||
            insertionRuns.Length != 1 ||
            !DocxLiteralTextSpanCodec.SameRunProperties(deletionRuns.Concat(insertionRuns)) ||
            deletedText.Length == 0 || insertedText.Length == 0 ||
            string.IsNullOrWhiteSpace(deletion.Id?.Value) || string.IsNullOrWhiteSpace(insertion.Id?.Value) ||
            deletion.Id?.Value == insertion.Id?.Value ||
            string.IsNullOrWhiteSpace(deletion.Author?.Value) || string.IsNullOrWhiteSpace(insertion.Author?.Value) ||
            deletion.Author?.Value != insertion.Author?.Value || deletion.Date?.Value != insertion.Date?.Value)
            return false;

        replacement = new DocxInlineTrackedReplacement(
            deletion,
            insertion,
            deletedText,
            insertedText,
            checked((uint)deletionRuns.Length),
            checked((uint)insertionRuns.Length));
        return true;
    }

    private static DocxResolvedTrackedReplacementTarget ResolveTarget(
        W.Body body,
        DocumentTrackedReplacementRequest request)
    {
        var selector = request.Target;
        var blockIndex = selector?.BlockIndex ?? request.TargetBlockIndex;
        if (selector is not null && selector.BlockIndex != request.TargetBlockIndex)
            throw new CodecException(
                "document_tracked_replacement_target_mismatch",
                "DOCX tracked replacement structured target block_index must match compatibility target_block_index.",
                DocumentPath);

        var blocks = body.ChildElements.Where(element => element is not W.SectionProperties).ToArray();
        if (blockIndex >= blocks.Length)
            throw new CodecException(
                "document_tracked_replacement_target_not_found",
                $"DOCX tracked replacement target block {blockIndex} is outside the {blocks.Length}-block document body.",
                DocumentPath);

        if (selector is null || selector.LocationCase == DocumentTrackedReplacementTarget.LocationOneofCase.BodyParagraph)
        {
            if (blocks[blockIndex] is not W.Paragraph paragraph)
                throw new CodecException(
                    "unsupported_document_tracked_replacement_target",
                    $"DOCX tracked replacement target block {blockIndex} is not a direct body paragraph.",
                    DocumentPath);
            return new DocxResolvedTrackedReplacementTarget(
                paragraph,
                BodyIndex(body, paragraph),
                new DocumentTrackedReplacementTarget
                {
                    BlockIndex = blockIndex,
                    BodyParagraph = new DocumentTrackedReplacementBodyParagraph(),
                });
        }

        if (selector.LocationCase != DocumentTrackedReplacementTarget.LocationOneofCase.TableCell ||
            selector.TableCell is null)
            throw new CodecException(
                "invalid_document_tracked_replacement_target",
                "DOCX tracked replacement target must select a body paragraph or table cell.",
                DocumentPath);
        if (blocks[blockIndex] is not W.Table table)
            throw new CodecException(
                "unsupported_document_tracked_replacement_target",
                $"DOCX tracked replacement target block {blockIndex} is not a direct body table.",
                DocumentPath);
        if (table.ChildElements.Any(child => child is not W.TableProperties and not W.TableGrid and not W.TableRow))
            throw Unsupported("Target table must contain only direct rows plus bounded table properties and grid metadata.");

        var rows = table.Elements<W.TableRow>().ToArray();
        if (selector.TableCell.Row >= rows.Length)
            throw TargetNotFound($"Target table row {selector.TableCell.Row} is outside the {rows.Length}-row physical table.");
        var rowIndex = (int)selector.TableCell.Row;
        var row = rows[rowIndex];
        if (row.ChildElements.Any(child => child is not W.TableRowProperties and not W.TableCell))
            throw Unsupported($"Target table row {rowIndex} contains unsupported native children.");
        var cells = row.Elements<W.TableCell>().ToArray();
        if (selector.TableCell.Column >= cells.Length)
            throw TargetNotFound($"Target table cell {rowIndex},{selector.TableCell.Column} is outside the {cells.Length}-cell physical row.");
        var columnIndex = (int)selector.TableCell.Column;

        var geometry = DocxTableGeometry.Read(table, out var validGeometry);
        if (!validGeometry || rowIndex >= geometry.Rows.Count || columnIndex >= geometry.Rows[rowIndex].RichCells.Count)
            throw Unsupported("Target table does not have a stable bounded physical grid.");
        var semanticCell = geometry.Rows[rowIndex].RichCells[columnIndex];
        if (semanticCell.VerticalMerge == DocumentTableVerticalMerge.Continue)
            throw Unsupported($"Target table cell {rowIndex},{columnIndex} is a vertical-merge continuation and has no independent editable text owner.");

        var cell = cells[columnIndex];
        if (cell.ChildElements.Any(child => child is not W.TableCellProperties and not W.Paragraph))
            throw Unsupported($"Target table cell {rowIndex},{columnIndex} contains nested tables, controls, or unsupported native children.");
        var paragraphs = cell.Elements<W.Paragraph>().ToArray();
        if (paragraphs.Length != 1)
            throw Unsupported($"Target table cell {rowIndex},{columnIndex} must contain exactly one direct paragraph; found {paragraphs.Length}.");

        return new DocxResolvedTrackedReplacementTarget(
            paragraphs[0],
            BodyIndex(body, table),
            selector.Clone());
    }

    private static DocxLiteralTextSpan FindTarget(W.Paragraph paragraph, string expectedText, string search)
    {
        var children = paragraph.ChildElements.Where(child => child is not W.ParagraphProperties).ToArray();
        if (children.Length == 0 || children.Any(child => child is not W.Run))
            throw Unsupported("Target paragraph must contain only direct ordinary text runs and no fields, hyperlinks, content controls, drawings, or existing revisions.");

        var runs = children.Cast<W.Run>().ToArray();
        foreach (var run in runs)
        {
            if (!TryReadTextRun(run, deleted: false, out _) || run.Elements<W.Text>().Count() != 1)
                throw Unsupported("Target paragraph must contain exactly one plain w:t node in each direct run.");
        }
        var resolution = DocxLiteralTextSpanCodec.Resolve(paragraph, expectedText, search);
        return resolution.Status switch
        {
            DocxLiteralTextSpanStatus.Success when resolution.Span is not null => resolution.Span,
            DocxLiteralTextSpanStatus.TextMismatch => throw new CodecException(
                "document_tracked_replacement_text_mismatch",
                "DOCX tracked replacement expected_paragraph_text does not match the exact target paragraph.",
                DocumentPath),
            DocxLiteralTextSpanStatus.MatchNotUnique => throw new CodecException(
                "document_tracked_replacement_match_not_unique",
                $"DOCX tracked replacement requires exactly one literal match in the target paragraph; found {resolution.MatchCount}.",
                DocumentPath),
            DocxLiteralTextSpanStatus.RunPropertyMismatch => throw new CodecException(
                "document_tracked_replacement_cross_run_format_mismatch",
                "DOCX tracked replacement cannot infer one insertion format from a literal spanning runs with different w:rPr markup.",
                DocumentPath),
            DocxLiteralTextSpanStatus.EmptyRunGap => throw Unsupported("A fragmented literal match cannot cross an empty native run."),
            _ => throw new CodecException(
                "document_tracked_replacement_verification_failed",
                $"DOCX tracked replacement could not map the unique literal to native run segments: {DocxLiteralTextSpanCodec.FailureDescription(resolution.Status)}.",
                DocumentPath),
        };
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

    private static bool TryReadRevisionRuns(
        OpenXmlCompositeElement revision,
        bool deleted,
        out string text,
        out W.Run[] runs)
    {
        text = string.Empty;
        runs = revision.Elements<W.Run>().ToArray();
        if (runs.Length == 0 || revision.ChildElements.Any(child => child is not W.Run)) return false;
        var values = new string[runs.Length];
        for (var index = 0; index < runs.Length; index++)
        {
            if (!TryReadTextRun(runs[index], deleted, out values[index]) || values[index].Length == 0) return false;
        }
        text = string.Concat(values);
        return true;
    }

    private static bool TryReadTextRun(W.Run run, bool deleted, out string text)
    {
        text = string.Empty;
        if (run.Elements<W.RunProperties>().Count() > 1 ||
            run.ChildElements.Any(child => child is not W.RunProperties &&
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
        if (request.Target is not null &&
            request.Target.LocationCase is not DocumentTrackedReplacementTarget.LocationOneofCase.BodyParagraph and
                not DocumentTrackedReplacementTarget.LocationOneofCase.TableCell)
            throw new CodecException(
                "invalid_document_tracked_replacement_target",
                "DOCX tracked replacement target must select a body paragraph or table cell.");
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

    private static uint BodyIndex(W.Body body, OpenXmlElement element)
    {
        for (var index = 0; index < body.ChildElements.Count; index++)
            if (ReferenceEquals(body.ChildElements[index], element)) return checked((uint)index);
        throw new CodecException("document_tracked_replacement_target_not_found", "DOCX tracked replacement target owner is not in the document body.", DocumentPath);
    }

    private static bool SameTarget(DocumentTrackedReplacementTarget left, DocumentTrackedReplacementTarget right) =>
        left.BlockIndex == right.BlockIndex &&
        left.LocationCase == right.LocationCase &&
        (left.LocationCase != DocumentTrackedReplacementTarget.LocationOneofCase.TableCell ||
         left.TableCell.Row == right.TableCell.Row && left.TableCell.Column == right.TableCell.Column);

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
    private static CodecException TargetNotFound(string message) => new("document_tracked_replacement_target_not_found", message, DocumentPath);
    private static CodecException Unsupported(string message) => new("unsupported_document_tracked_replacement_topology", message, DocumentPath);
}
