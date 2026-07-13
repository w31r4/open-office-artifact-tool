using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenOffice.OpenXmlCodec;

internal sealed record DocxImportResult(ArtifactEnvelope Artifact, IReadOnlyList<Diagnostic> Diagnostics);
internal sealed record DocxExportResult(byte[] File, IReadOnlyList<Diagnostic> Diagnostics);

internal static class DocxCodec
{
    internal static DocxImportResult Import(byte[] bytes, EffectiveCodecLimits limits)
    {
        var opaque = PackageGuards.ValidateAndCollectOpaque(bytes, limits, OpcPackageProfile.Docx);
        var diagnostics = new List<Diagnostic>();
        var opaqueCount = opaque.Parts.Count + opaque.PackageRelationships.Count;
        if (opaqueCount > 0)
            diagnostics.Add(CodecProtocol.Warning(
                "opaque_content_retained",
                $"Retained {opaqueCount} unsupported OPC parts or relationships with a hash-bound source package snapshot for loss-aware export.",
                opaque.Parts.FirstOrDefault()?.Path ?? opaque.PackageRelationships.FirstOrDefault()?.SourcePath));

        using var stream = new MemoryStream(bytes, writable: false);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var mainPart = package.MainDocumentPart ??
            throw new CodecException("missing_document_part", "DOCX package has no Main Document part.", "word/document.xml");
        var body = mainPart.Document?.Body ??
            throw new CodecException("missing_document_body", "DOCX package has no document body.", "word/document.xml");

        var document = new DocumentArtifact { Id = "document/1", Name = "Imported document" };
        ulong semanticItems = 0;
        var ordinal = 0;
        for (var bodyIndex = 0; bodyIndex < body.ChildElements.Count; bodyIndex++)
        {
            var element = body.ChildElements[bodyIndex];
            if (element is W.SectionProperties) continue;
            var block = ReadBodyBlock(element, ordinal++, checked((uint)bodyIndex), ref semanticItems, limits);
            document.Blocks.Add(block);
        }

        var envelope = new ArtifactEnvelope
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Family = ArtifactFamily.Document,
            Document = document,
            OpaqueOpc = opaque,
            Source = new SourceIdentity
            {
                Format = "docx",
                PackageSha256 = Hash(bytes),
                Producer = "open-office-artifact-tool/OpenXmlWasm",
            },
        };
        envelope.Diagnostics.Add(diagnostics);
        return new DocxImportResult(envelope, diagnostics);
    }

    internal static DocxExportResult Export(ArtifactEnvelope envelope, EffectiveCodecLimits limits, bool allowLossy)
    {
        ValidateEnvelope(envelope, limits);
        var opaqueCount = (envelope.OpaqueOpc?.Parts.Count ?? 0) +
                          (envelope.OpaqueOpc?.PackageRelationships.Count ?? 0);
        if (envelope.OpaqueOpc?.SourcePackage is { Data.IsEmpty: false })
            return ExportPreservingSource(envelope, limits, opaqueCount);
        if (opaqueCount > 0 && !allowLossy)
            throw new CodecException(
                "opaque_content_requires_preservation",
                "Document contains opaque OPC parts or relationships but its validated source package snapshot is unavailable; pass allow_lossy only when discarding them is intentional.");

        var diagnostics = new List<Diagnostic>();
        if (opaqueCount > 0)
            diagnostics.Add(CodecProtocol.Warning(
                "opaque_content_discarded",
                $"Discarded {opaqueCount} opaque OPC parts or relationships under explicit allow_lossy policy."));

        using var stream = new MemoryStream();
        using (var package = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: true))
        {
            var mainPart = package.AddMainDocumentPart();
            var body = new W.Body();
            foreach (var block in envelope.Document.Blocks) body.Append(BuildBlock(block));
            body.Append(new W.SectionProperties());
            mainPart.Document = new W.Document(body);
            mainPart.Document.Save();
        }

        var bytes = stream.ToArray();
        ValidateOutputBudget(bytes, limits);
        ValidateOffice2021(bytes);
        return new DocxExportResult(bytes, diagnostics);
    }

    private static DocxExportResult ExportPreservingSource(ArtifactEnvelope envelope, EffectiveCodecLimits limits, int opaqueCount)
    {
        var sourceBytes = PackageGuards.ValidateSourcePackage(
            envelope.OpaqueOpc,
            envelope.Source,
            limits,
            OpcPackageProfile.Docx);
        using var stream = new MemoryStream();
        stream.Write(sourceBytes);
        stream.Position = 0;
        using (var package = WordprocessingDocument.Open(stream, isEditable: true, new OpenSettings { AutoSave = true }))
        {
            var mainPart = package.MainDocumentPart ??
                throw new CodecException("missing_document_part", "DOCX package has no Main Document part.", "word/document.xml");
            var body = mainPart.Document?.Body ??
                throw new CodecException("missing_document_body", "DOCX package has no document body.", "word/document.xml");
            var sourceElements = body.ChildElements.Where(element => element is not W.SectionProperties).ToArray();
            if (sourceElements.Length != envelope.Document.Blocks.Count)
                throw new CodecException(
                    "document_topology_changed",
                    $"Source-preserving DOCX export requires the original {sourceElements.Length}-block body topology; the artifact contains {envelope.Document.Blocks.Count} blocks.",
                    "word/document.xml");

            ulong semanticItems = 0;
            for (var ordinal = 0; ordinal < envelope.Document.Blocks.Count; ordinal++)
            {
                var block = envelope.Document.Blocks[ordinal];
                var binding = block.Source ?? throw new CodecException(
                    "missing_document_source_binding",
                    $"Document block {ordinal} is missing its source binding.",
                    "word/document.xml");
                if (binding.BodyIndex >= body.ChildElements.Count || !ReferenceEquals(body.ChildElements[(int)binding.BodyIndex], sourceElements[ordinal]))
                    throw new CodecException(
                        "document_source_binding_mismatch",
                        $"Document block {ordinal} does not point to its original body element.",
                        "word/document.xml");

                var element = sourceElements[ordinal];
                if (!HashElement(element).Equals(binding.ElementSha256, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "document_source_element_mismatch",
                        $"Document block {ordinal} no longer matches its source element hash.",
                        "word/document.xml");

                var original = ReadBodyBlock(element, ordinal, binding.BodyIndex, ref semanticItems, limits);
                if (!SemanticHash(original).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "document_source_semantics_mismatch",
                        $"Document block {ordinal} source semantics do not match its binding.",
                        "word/document.xml");

                if (SemanticHash(block).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase)) continue;
                if (!binding.Editable || block.ContentCase != original.ContentCase || block.ContentCase == DocumentBlock.ContentOneofCase.Opaque)
                    throw new CodecException(
                        "unsupported_document_edit",
                        $"Document block {ordinal} contains WordprocessingML that is preserved but not yet safely editable by this codec slice.",
                        "word/document.xml");

                element.InsertBeforeSelf(BuildBlock(block));
                element.Remove();
            }
            mainPart.Document!.Save();
        }

        var bytes = stream.ToArray();
        ValidateOutputBudget(bytes, limits);
        ValidateOffice2021(bytes);
        var outputOpaque = PackageGuards.ValidateAndCollectOpaque(bytes, limits, OpcPackageProfile.Docx, includeSourcePackage: false);
        PackageGuards.AssertOpaqueGraphMatches(envelope.OpaqueOpc, outputOpaque, "opaque_content_not_preserved");
        var diagnostics = new List<Diagnostic>();
        if (opaqueCount > 0)
            diagnostics.Add(CodecProtocol.Warning(
                "opaque_content_preserved",
                $"Preserved {opaqueCount} opaque OPC parts or relationships while updating modeled document content."));
        return new DocxExportResult(bytes, diagnostics);
    }

    private static DocumentBlock ReadBodyBlock(
        OpenXmlElement element,
        int ordinal,
        uint bodyIndex,
        ref ulong semanticItems,
        EffectiveCodecLimits limits)
    {
        var block = new DocumentBlock { Id = $"document/block/{ordinal + 1}" };
        var editable = false;
        switch (element)
        {
            case W.Paragraph paragraph:
                editable = IsSimpleParagraph(paragraph);
                block.StyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? string.Empty;
                var paragraphArtifact = new DocumentParagraph { Text = DescendantText(paragraph) };
                if (editable)
                {
                    foreach (var run in paragraph.Elements<W.Run>()) paragraphArtifact.Runs.Add(ReadRun(run));
                }
                block.Paragraph = paragraphArtifact;
                semanticItems += checked((ulong)Math.Max(1, paragraphArtifact.Runs.Count));
                break;
            case W.Table table:
                block.StyleId = table.TableProperties?.TableStyle?.Val?.Value ?? string.Empty;
                var tableArtifact = new DocumentTable();
                foreach (var row in table.Elements<W.TableRow>())
                {
                    var targetRow = new DocumentTableRow();
                    foreach (var cell in row.Elements<W.TableCell>())
                    {
                        targetRow.Cells.Add(DescendantText(cell));
                        semanticItems++;
                    }
                    tableArtifact.Rows.Add(targetRow);
                }
                block.Table = tableArtifact;
                break;
            default:
                block.Opaque = new DocumentOpaqueBlock
                {
                    ElementName = element.LocalName,
                    Text = element.InnerText ?? string.Empty,
                };
                semanticItems++;
                break;
        }
        if (semanticItems > limits.MaxCells)
            throw new CodecException(
                "document_item_budget_exceeded",
                $"DOCX document exceeds max_cells semantic-item budget ({limits.MaxCells}).",
                "word/document.xml");

        block.Source = new DocumentSourceBinding
        {
            BodyIndex = bodyIndex,
            ElementSha256 = HashElement(element),
            Editable = editable,
        };
        block.Source.SemanticSha256 = SemanticHash(block);
        return block;
    }

    private static DocumentRun ReadRun(W.Run run)
    {
        var properties = run.RunProperties;
        return new DocumentRun
        {
            Text = DescendantText(run),
            StyleId = properties?.RunStyle?.Val?.Value ?? string.Empty,
            Bold = IsOn(properties?.Bold),
            Italic = IsOn(properties?.Italic),
            Underline = IsUnderline(properties?.Underline),
        };
    }

    private static bool IsSimpleParagraph(W.Paragraph paragraph)
    {
        if (paragraph.ChildElements.Any(child => child is not W.ParagraphProperties and not W.Run)) return false;
        if (paragraph.ParagraphProperties?.ChildElements.Any(child => child is not W.ParagraphStyleId) == true) return false;
        return paragraph.Elements<W.Run>().All(run =>
            run.ChildElements.All(child => child is W.RunProperties or W.Text) &&
            (run.RunProperties?.ChildElements.All(child =>
                child is W.RunStyle or W.Bold or W.Italic or W.Underline) ?? true));
    }

    private static OpenXmlElement BuildBlock(DocumentBlock block) => block.ContentCase switch
    {
        DocumentBlock.ContentOneofCase.Paragraph => BuildParagraph(block),
        DocumentBlock.ContentOneofCase.Table => BuildTable(block),
        DocumentBlock.ContentOneofCase.Opaque => throw new CodecException(
            "unsupported_document_block",
            $"Opaque document block {block.Id} requires its validated source package and cannot be authored from scratch."),
        _ => throw new CodecException("missing_document_block_content", $"Document block {block.Id} has no content."),
    };

    private static W.Paragraph BuildParagraph(DocumentBlock block)
    {
        var paragraph = new W.Paragraph();
        if (!string.IsNullOrWhiteSpace(block.StyleId))
            paragraph.ParagraphProperties = new W.ParagraphProperties(new W.ParagraphStyleId { Val = block.StyleId });
        if (block.Paragraph.Runs.Count == 0)
        {
            if (block.Paragraph.Text.Length > 0) paragraph.Append(new W.Run(Text(block.Paragraph.Text)));
            return paragraph;
        }
        foreach (var source in block.Paragraph.Runs)
        {
            var run = new W.Run();
            var properties = new W.RunProperties();
            if (!string.IsNullOrWhiteSpace(source.StyleId)) properties.Append(new W.RunStyle { Val = source.StyleId });
            if (source.Bold) properties.Append(new W.Bold());
            if (source.Italic) properties.Append(new W.Italic());
            if (source.Underline) properties.Append(new W.Underline { Val = W.UnderlineValues.Single });
            if (properties.ChildElements.Count > 0) run.Append(properties);
            run.Append(Text(source.Text));
            paragraph.Append(run);
        }
        return paragraph;
    }

    private static W.Table BuildTable(DocumentBlock block)
    {
        var table = new W.Table();
        if (!string.IsNullOrWhiteSpace(block.StyleId))
            table.Append(new W.TableProperties(new W.TableStyle { Val = block.StyleId }));
        var columns = block.Table.Rows.Count == 0 ? 1 : Math.Max(1, block.Table.Rows.Max(row => row.Cells.Count));
        var grid = new W.TableGrid();
        for (var column = 0; column < columns; column++) grid.Append(new W.GridColumn());
        table.Append(grid);
        foreach (var sourceRow in block.Table.Rows)
        {
            var row = new W.TableRow();
            foreach (var value in sourceRow.Cells)
                row.Append(new W.TableCell(new W.Paragraph(new W.Run(Text(value)))));
            table.Append(row);
        }
        return table;
    }

    private static W.Text Text(string value) => new(value)
    {
        Space = value.Length != value.Trim().Length ? SpaceProcessingModeValues.Preserve : null,
    };

    private static string DescendantText(OpenXmlElement element) =>
        string.Concat(element.Descendants<W.Text>().Select(text => text.Text));

    private static bool IsOn(W.OnOffType? value) => value is not null && value.Val?.Value != false;

    private static bool IsUnderline(W.Underline? value)
    {
        var underline = value?.Val?.Value;
        return underline is not null && !underline.Equals(W.UnderlineValues.None);
    }

    private static string SemanticHash(DocumentBlock block)
    {
        var semantic = block.Clone();
        semantic.Id = string.Empty;
        semantic.Name = string.Empty;
        semantic.Source = null;
        return Hash(semantic.ToByteArray());
    }

    private static string HashElement(OpenXmlElement element) => Hash(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static void ValidateEnvelope(ArtifactEnvelope envelope, EffectiveCodecLimits limits)
    {
        if (envelope.ProtocolVersion != CodecProtocol.ProtocolVersion)
            throw new CodecException("unsupported_artifact_version", $"Artifact protocol version {envelope.ProtocolVersion} is unsupported.");
        if (envelope.Family != ArtifactFamily.Document || envelope.PayloadCase != ArtifactEnvelope.PayloadOneofCase.Document)
            throw new CodecException("invalid_document_artifact", "Artifact envelope does not contain a document payload.");
        if ((ulong)envelope.Document.Blocks.Count > limits.MaxCells)
            throw new CodecException("document_item_budget_exceeded", $"Document has {envelope.Document.Blocks.Count} blocks and exceeds max_cells ({limits.MaxCells}).");

        ulong semanticItems = 0;
        foreach (var block in envelope.Document.Blocks)
        {
            switch (block.ContentCase)
            {
                case DocumentBlock.ContentOneofCase.Paragraph:
                    if (block.Paragraph.Runs.Count > 0 && block.Paragraph.Text != string.Concat(block.Paragraph.Runs.Select(run => run.Text)))
                        throw new CodecException("inconsistent_document_text", $"Document paragraph {block.Id} text does not match its runs.");
                    semanticItems += checked((ulong)Math.Max(1, block.Paragraph.Runs.Count));
                    break;
                case DocumentBlock.ContentOneofCase.Table:
                    semanticItems += checked((ulong)block.Table.Rows.Sum(row => row.Cells.Count));
                    break;
                case DocumentBlock.ContentOneofCase.Opaque:
                    semanticItems++;
                    break;
                default:
                    throw new CodecException("missing_document_block_content", $"Document block {block.Id} has no content.");
            }
            if (semanticItems > limits.MaxCells)
                throw new CodecException("document_item_budget_exceeded", $"Document exceeds max_cells semantic-item budget ({limits.MaxCells}).");
        }
    }

    private static void ValidateOutputBudget(byte[] bytes, EffectiveCodecLimits limits)
    {
        if ((ulong)bytes.LongLength > limits.MaxInputBytes)
            throw new CodecException("output_budget_exceeded", $"Generated DOCX has {bytes.LongLength} bytes and exceeds max_input_bytes ({limits.MaxInputBytes}).");
    }

    private static void ValidateOffice2021(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package).Take(8).ToArray();
        if (errors.Length == 0) return;
        var detail = string.Join("; ", errors.Select(error => $"{error.Path?.XPath ?? error.Part?.Uri.ToString() ?? "package"}: {error.Description}"));
        throw new CodecException("openxml_validation_failed", $"Generated DOCX is not valid Office 2021 Open XML: {detail}");
    }
}
