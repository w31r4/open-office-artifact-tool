using System.IO.Compression;
using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

internal sealed record DocxRevisionFinalizationOutput(
    byte[] File,
    DocumentRevisionFinalizationResult Result,
    IReadOnlyList<Diagnostic> Diagnostics);

// Finalizes the bounded whole-paragraph profile plus the exact adjacent
// in-paragraph deletion/insertion pair authored by DocxTrackedReplacementCodec.
// The pair may live in a direct body paragraph or one structurally bounded
// direct table-cell paragraph. Every other topology fails before bytes are
// written.
internal static class DocxRevisionFinalizationCodec
{
    private const string DocumentPath = "word/document.xml";
    private const string SettingsPath = "word/settings.xml";

    internal static DocxRevisionFinalizationOutput Finalize(
        byte[] sourceBytes,
        DocumentRevisionFinalizationRequest request,
        EffectiveCodecLimits limits)
    {
        if (request.Mode is not DocumentRevisionFinalizationMode.Accept and not DocumentRevisionFinalizationMode.Reject)
            throw new CodecException(
                "invalid_document_revision_finalization",
                "DOCX revision finalization mode must be accept or reject.");

        var sourceHash = Hash(sourceBytes);
        if (!IsSha256(request.ExpectedSourceSha256) ||
            !sourceHash.Equals(request.ExpectedSourceSha256, StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "document_source_hash_mismatch",
                "DOCX revision finalization requires expected_source_sha256 to match the exact input bytes.");

        var sourceOpaque = PackageGuards.ValidateAndCollectOpaque(sourceBytes, limits, OpcPackageProfile.Docx, includeSourcePackage: false);
        var revisionInventory = DocxRevisionMarkup.Inventory(sourceBytes);
        var outsideMainDocument = revisionInventory
            .Where(item => !item.Key.Equals(DocumentPath, StringComparison.OrdinalIgnoreCase) && item.Value > 0)
            .Select(item => item.Key)
            .OrderBy(item => item, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (outsideMainDocument.Length > 0)
            throw new CodecException(
                "unsupported_document_revision_scope",
                $"Revision markup outside {DocumentPath} is not supported by this bounded finalizer: {string.Join(", ", outsideMainDocument)}.",
                outsideMainDocument[0]);

        var mainRevisionCount = revisionInventory.GetValueOrDefault(DocumentPath);
        if (mainRevisionCount == 0)
            throw new CodecException(
                "document_revisions_not_found",
                "DOCX revision finalization requires at least one supported tracked insertion or deletion.",
                DocumentPath);
        if ((ulong)mainRevisionCount > limits.MaxCells)
            throw new CodecException(
                "document_item_budget_exceeded",
                $"DOCX contains {mainRevisionCount} revision markers and exceeds max_cells ({limits.MaxCells}).",
                DocumentPath);

        using var stream = new MemoryStream();
        stream.Write(sourceBytes);
        stream.Position = 0;
        uint insertionCount = 0;
        uint deletionCount = 0;
        bool trackingBefore;
        bool trackingAfter;
        using (var package = WordprocessingDocument.Open(stream, isEditable: true, new OpenSettings { AutoSave = false }))
        {
            var mainPart = package.MainDocumentPart ??
                throw new CodecException("missing_document_part", "DOCX package has no Main Document part.", DocumentPath);
            var body = mainPart.Document?.Body ??
                throw new CodecException("missing_document_body", "DOCX package has no document body.", DocumentPath);
            var supported = new List<OpenXmlCompositeElement>();
            foreach (var paragraph in body.Descendants<W.Paragraph>())
            {
                var directRevisions = paragraph.ChildElements
                    .Where(child => child is W.InsertedRun or W.DeletedRun)
                    .Cast<OpenXmlCompositeElement>()
                    .ToArray();
                if (directRevisions.Length == 0) continue;
                if (paragraph.Parent is W.Body &&
                    DocxTrackedChangeCodec.TryRead(paragraph, out _, out _, out var editable) && editable)
                {
                    supported.Add(directRevisions.Single());
                    continue;
                }
                var boundedOwner = paragraph.Parent is W.Body || IsBoundedTableCellParagraph(paragraph);
                if (boundedOwner && DocxTrackedReplacementCodec.TryReadFinalizable(paragraph, out var replacement))
                {
                    supported.Add(replacement.Deletion);
                    supported.Add(replacement.Insertion);
                }
            }
            if (supported.Count != mainRevisionCount)
                throw new CodecException(
                    "unsupported_document_revision_topology",
                    "DOCX contains unsupported revision markup. This bounded finalizer accepts direct whole-paragraph one-run w:ins/w:del blocks in the body, or one adjacent direct-run w:del + w:ins replacement pair in a direct body paragraph or one bounded direct table-cell paragraph; mixed, nested, moved, property-level, non-body-story, irregular-table, or malformed graphs fail closed.",
                    DocumentPath);

            foreach (var wrapper in supported)
            {
                switch (wrapper)
                {
                    case W.InsertedRun insertion:
                        insertionCount++;
                        if (request.Mode == DocumentRevisionFinalizationMode.Accept) Unwrap(insertion, deleted: false);
                        else insertion.Remove();
                        break;
                    case W.DeletedRun deletion:
                        deletionCount++;
                        if (request.Mode == DocumentRevisionFinalizationMode.Reject) Unwrap(deletion, deleted: true);
                        else deletion.Remove();
                        break;
                }
            }
            mainPart.Document!.Save();

            var settings = mainPart.DocumentSettingsPart?.Settings;
            trackingBefore = Enabled(settings?.GetFirstChild<W.TrackRevisions>());
            if (!request.KeepTracking && trackingBefore)
            {
                settings!.RemoveAllChildren<W.TrackRevisions>();
                settings.Save();
            }
            trackingAfter = request.KeepTracking && trackingBefore;
        }

        var outputBytes = stream.ToArray();
        DocxCodec.ValidateOutputBudget(outputBytes, limits);
        var retainedValidationErrorCount = DocxCodec.ValidateOffice2021AgainstSource(sourceBytes, outputBytes);
        if (DocxRevisionMarkup.Inventory(outputBytes).Values.Sum() != 0)
            throw new CodecException(
                "document_revision_finalization_incomplete",
                "DOCX still contains revision markup after bounded finalization.",
                DocumentPath);

        var changedParts = ChangedParts(sourceBytes, outputBytes);
        if (!changedParts.Contains(DocumentPath, StringComparer.OrdinalIgnoreCase) ||
            changedParts.Any(path =>
                !path.Equals(DocumentPath, StringComparison.OrdinalIgnoreCase) &&
                !path.Equals(SettingsPath, StringComparison.OrdinalIgnoreCase)))
            throw new CodecException(
                "document_revision_scope_violation",
                $"Revision finalization changed unexpected OPC parts: {string.Join(", ", changedParts)}.");
        if (changedParts.Contains(SettingsPath, StringComparer.OrdinalIgnoreCase) && (request.KeepTracking || !trackingBefore))
            throw new CodecException(
                "document_revision_scope_violation",
                "Revision finalization changed word/settings.xml without removing an enabled trackRevisions setting.",
                SettingsPath);

        var outputOpaque = PackageGuards.ValidateAndCollectOpaque(outputBytes, limits, OpcPackageProfile.Docx, includeSourcePackage: false);
        PackageGuards.AssertOpaqueGraphMatches(
            sourceOpaque,
            outputOpaque,
            "opaque_content_not_preserved",
            ignorePart: changedParts.Contains(SettingsPath, StringComparer.OrdinalIgnoreCase)
                ? part => part.Path.Equals(SettingsPath, StringComparison.OrdinalIgnoreCase)
                : null);

        var result = new DocumentRevisionFinalizationResult
        {
            Mode = request.Mode,
            SourceSha256 = sourceHash,
            OutputSha256 = Hash(outputBytes),
            InsertionCount = insertionCount,
            DeletionCount = deletionCount,
            TrackingBefore = trackingBefore,
            TrackingAfter = trackingAfter,
        };
        result.ChangedParts.Add(changedParts);
        var diagnostics = new List<Diagnostic>();
        if (retainedValidationErrorCount > 0)
            diagnostics.Add(CodecProtocol.Warning(
                "source_openxml_validation_warnings_preserved",
                $"Preserved {retainedValidationErrorCount} pre-existing Office 2021 validation warning(s) from the source package; revision finalization introduced none."));
        return new DocxRevisionFinalizationOutput(outputBytes, result, diagnostics);
    }

    private static void Unwrap(OpenXmlCompositeElement revision, bool deleted)
    {
        foreach (var sourceRun in revision.Elements<W.Run>().ToArray())
        {
            var run = (W.Run)sourceRun.CloneNode(true);
            if (deleted)
            {
                foreach (var deletedText in run.Descendants<W.DeletedText>().ToArray())
                {
                    var text = new W.Text(deletedText.Text) { Space = deletedText.Space };
                    deletedText.InsertAfterSelf(text);
                    deletedText.Remove();
                }
            }
            revision.InsertBeforeSelf(run);
        }
        revision.Remove();
    }

    private static bool IsBoundedTableCellParagraph(W.Paragraph paragraph)
    {
        if (paragraph.Parent is not W.TableCell cell ||
            cell.Parent is not W.TableRow row ||
            row.Parent is not W.Table table ||
            table.Parent is not W.Body)
            return false;
        if (cell.ChildElements.Any(child => child is not W.TableCellProperties and not W.Paragraph) ||
            cell.Elements<W.Paragraph>().Count() != 1 ||
            row.ChildElements.Any(child => child is not W.TableRowProperties and not W.TableCell) ||
            table.ChildElements.Any(child => child is not W.TableProperties and not W.TableGrid and not W.TableRow))
            return false;

        var geometry = DocxTableGeometry.Read(table, out var validGeometry);
        if (!validGeometry) return false;
        var rowIndex = table.Elements<W.TableRow>().TakeWhile(candidate => !ReferenceEquals(candidate, row)).Count();
        var columnIndex = row.Elements<W.TableCell>().TakeWhile(candidate => !ReferenceEquals(candidate, cell)).Count();
        return rowIndex < geometry.Rows.Count &&
               columnIndex < geometry.Rows[rowIndex].RichCells.Count &&
               geometry.Rows[rowIndex].RichCells[columnIndex].VerticalMerge != DocumentTableVerticalMerge.Continue;
    }

    private static string[] ChangedParts(byte[] sourceBytes, byte[] outputBytes)
    {
        var source = PartHashes(sourceBytes);
        var output = PartHashes(outputBytes);
        if (!source.Keys.OrderBy(item => item, StringComparer.OrdinalIgnoreCase)
                .SequenceEqual(output.Keys.OrderBy(item => item, StringComparer.OrdinalIgnoreCase), StringComparer.OrdinalIgnoreCase))
            throw new CodecException(
                "document_revision_scope_violation",
                "Revision finalization changed the OPC part inventory.");
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

    private static bool Enabled(W.OnOffType? value) => value is not null && value.Val?.Value != false;
    private static bool IsSha256(string value) => value.Length == 64 && value.All(Uri.IsHexDigit);
    private static string Hash(ReadOnlySpan<byte> bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
