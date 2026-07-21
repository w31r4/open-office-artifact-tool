using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OpenChestnut.Codec;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OpenChestnut.Codec.Tests;

public sealed class DocxTrackedReplacementCodecTests
{
    [Fact]
    public void AddsAuditedNativeReplacementAndFinalizesAcceptOrReject()
    {
        const string sourceText = "The draft budget assumes 30 days of cash buffer.";
        const string oldText = "30 days";
        const string newText = "45 days";
        var source = ExportDocx(Paragraph(sourceText, bold: true));
        var sourceHash = Hash(source);

        var added = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.AddDocxTrackedReplacement,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(source),
            TrackedReplacement = new DocumentTrackedReplacementRequest
            {
                ExpectedSourceSha256 = sourceHash,
                TargetBlockIndex = 0,
                ExpectedParagraphText = sourceText,
                Search = oldText,
                Replacement = newText,
                Author = "Budget reviewer",
                Date = "2026-07-21T09:30:00Z",
            },
        });

        Assert.True(added.Ok, Diagnostics(added));
        Assert.NotEmpty(added.File);
        var audit = Assert.IsType<DocumentTrackedReplacementResult>(added.TrackedReplacement);
        Assert.Equal(sourceHash, audit.SourceSha256);
        Assert.Equal(Hash(added.File.Span), audit.OutputSha256);
        Assert.Equal((uint)0, audit.TargetBlockIndex);
        Assert.Equal((uint)0, audit.TargetBodyIndex);
        Assert.Equal(DocumentTrackedReplacementTarget.LocationOneofCase.BodyParagraph, audit.Target.LocationCase);
        Assert.Equal((uint)0, audit.Target.BlockIndex);
        Assert.Equal(HashText(oldText), audit.DeletedTextSha256);
        Assert.Equal(HashText(newText), audit.InsertedTextSha256);
        Assert.Equal((uint)oldText.Length, audit.DeletedTextChars);
        Assert.Equal((uint)newText.Length, audit.InsertedTextChars);
        Assert.NotEqual(audit.DeletionNativeRevisionId, audit.InsertionNativeRevisionId);
        Assert.Equal(["word/document.xml"], audit.ChangedParts);

        using (var stream = new MemoryStream(added.File.ToByteArray(), writable: false))
        using (var package = WordprocessingDocument.Open(stream, isEditable: false))
        {
            var paragraph = package.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().Single();
            var deletion = paragraph.Elements<W.DeletedRun>().Single();
            var insertion = paragraph.Elements<W.InsertedRun>().Single();
            Assert.Equal(oldText, deletion.Descendants<W.DeletedText>().Single().Text);
            Assert.Equal(newText, insertion.Descendants<W.Text>().Single().Text);
            Assert.NotNull(deletion.Descendants<W.RunProperties>().Single().Bold);
            Assert.NotNull(insertion.Descendants<W.RunProperties>().Single().Bold);
            Assert.Equal(audit.DeletionNativeRevisionId, deletion.Id!.Value);
            Assert.Equal(audit.InsertionNativeRevisionId, insertion.Id!.Value);
            Assert.Equal(sourceText.Replace(oldText, newText), string.Concat(paragraph.Descendants<W.Text>().Select(value => value.Text)));
        }

        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = added.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var importedBlock = Assert.Single(imported.Artifact.Document.Blocks);
        Assert.Equal(DocumentBlock.ContentOneofCase.Paragraph, importedBlock.ContentCase);
        Assert.False(importedBlock.Source.Editable);
        Assert.Equal(sourceText.Replace(oldText, newText), importedBlock.Paragraph.Text);

        var accepted = Finalize(added.File.ToByteArray(), mode: DocumentRevisionFinalizationMode.Accept);
        Assert.Equal((uint)1, accepted.RevisionFinalization.InsertionCount);
        Assert.Equal((uint)1, accepted.RevisionFinalization.DeletionCount);
        Assert.Equal(sourceText.Replace(oldText, newText), BodyText(accepted.File.ToByteArray()));
        AssertNoRevisions(accepted.File.ToByteArray());

        var rejected = Finalize(added.File.ToByteArray(), mode: DocumentRevisionFinalizationMode.Reject);
        Assert.Equal(sourceText, BodyText(rejected.File.ToByteArray()));
        AssertNoRevisions(rejected.File.ToByteArray());
        Assert.Equal(sourceHash, Hash(source));
    }

    [Fact]
    public void AddsAndFinalizesTrackedReplacementInBoundedTableCell()
    {
        const string sourceText = "Payment is due in 30 days.";
        const string oldText = "30 days";
        const string newText = "45 days";
        var table = new DocumentBlock
        {
            Id = "table/terms",
            StyleId = "TableGrid",
            Table = new DocumentTable
            {
                GridColumns = 2,
                Formatting = new DocumentTableFormatting
                {
                    WidthDxa = 9000,
                    IndentDxa = 120,
                    CellMarginsDxa = new DocumentTableCellMargins { Top = 80, Bottom = 80, Start = 120, End = 120 },
                    BorderColor = "445566",
                    BorderSize = 8,
                    HeaderFill = "E2E8F0",
                },
            },
        };
        table.Table.Formatting.ColumnWidthsDxa.Add([2800U, 6200U]);
        var header = new DocumentTableRow();
        header.Cells.Add(["Term", "Value"]);
        var terms = new DocumentTableRow();
        terms.Cells.Add(["Payment", sourceText]);
        table.Table.Rows.Add([header, terms]);
        var source = ExportDocx(table);
        var sourceHash = Hash(source);

        var added = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.AddDocxTrackedReplacement,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(source),
            TrackedReplacement = new DocumentTrackedReplacementRequest
            {
                ExpectedSourceSha256 = sourceHash,
                TargetBlockIndex = 0,
                Target = new DocumentTrackedReplacementTarget
                {
                    BlockIndex = 0,
                    TableCell = new DocumentTrackedReplacementTableCell { Row = 1, Column = 1 },
                },
                ExpectedParagraphText = sourceText,
                Search = oldText,
                Replacement = newText,
                Author = "Contract reviewer",
                Date = "2026-07-21T11:00:00Z",
            },
        });

        Assert.True(added.Ok, Diagnostics(added));
        var audit = Assert.IsType<DocumentTrackedReplacementResult>(added.TrackedReplacement);
        Assert.Equal(DocumentTrackedReplacementTarget.LocationOneofCase.TableCell, audit.Target.LocationCase);
        Assert.Equal((uint)0, audit.Target.BlockIndex);
        Assert.Equal((uint)1, audit.Target.TableCell.Row);
        Assert.Equal((uint)1, audit.Target.TableCell.Column);
        Assert.Equal(["word/document.xml"], audit.ChangedParts);

        using (var stream = new MemoryStream(added.File.ToByteArray(), writable: false))
        using (var package = WordprocessingDocument.Open(stream, isEditable: false))
        {
            var nativeTable = package.MainDocumentPart!.Document!.Body!.Elements<W.Table>().Single();
            var cells = nativeTable.Elements<W.TableRow>().ElementAt(1).Elements<W.TableCell>().ToArray();
            Assert.Equal("Payment", string.Concat(cells[0].Descendants<W.Text>().Select(value => value.Text)));
            var paragraph = cells[1].Elements<W.Paragraph>().Single();
            Assert.Equal(oldText, paragraph.Elements<W.DeletedRun>().Single().Descendants<W.DeletedText>().Single().Text);
            Assert.Equal(newText, paragraph.Elements<W.InsertedRun>().Single().Descendants<W.Text>().Single().Text);
            Assert.Equal(sourceText.Replace(oldText, newText), string.Concat(paragraph.Descendants<W.Text>().Select(value => value.Text)));
        }

        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = added.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var importedTable = Assert.Single(imported.Artifact.Document.Blocks).Table;
        Assert.Equal(sourceText.Replace(oldText, newText), importedTable.Rows[1].Cells[1]);

        var accepted = Finalize(added.File.ToByteArray(), DocumentRevisionFinalizationMode.Accept);
        Assert.Equal(sourceText.Replace(oldText, newText), TableCellText(accepted.File.ToByteArray(), 1, 1));
        AssertNoRevisions(accepted.File.ToByteArray());

        var rejected = Finalize(added.File.ToByteArray(), DocumentRevisionFinalizationMode.Reject);
        Assert.Equal(sourceText, TableCellText(rejected.File.ToByteArray(), 1, 1));
        AssertNoRevisions(rejected.File.ToByteArray());
        Assert.Equal(sourceHash, Hash(source));
    }

    [Fact]
    public void FailsClosedForStaleAmbiguousCrossRunAndUnsupportedTargets()
    {
        const string repeated = "draft wording and draft wording";
        var repeatedSource = ExportDocx(Paragraph(repeated));
        AssertFailure(Add(repeatedSource, repeated, "draft", "final"), "document_tracked_replacement_match_not_unique");
        AssertFailure(Add(repeatedSource, repeated, "draft", "final", expectedHash: new string('0', 64)), "document_source_hash_mismatch");
        AssertFailure(Add(repeatedSource, "stale paragraph", "draft", "final"), "document_tracked_replacement_text_mismatch");
        AssertFailure(Add(repeatedSource, repeated, "draft", "final", targetBlockIndex: 9), "document_tracked_replacement_target_not_found");

        var crossRun = new DocumentBlock
        {
            Id = "paragraph/cross-run",
            StyleId = "Normal",
            Paragraph = new DocumentParagraph { Text = "cash buffer" },
        };
        crossRun.Paragraph.Runs.Add(new DocumentRun { Text = "cash " });
        crossRun.Paragraph.Runs.Add(new DocumentRun { Text = "buffer" });
        var crossRunSource = ExportDocx(crossRun);
        AssertFailure(Add(crossRunSource, "cash buffer", "cash buffer", "liquidity reserve"), "document_tracked_replacement_cross_run_match");

        var table = new DocumentBlock
        {
            Id = "table/1",
            StyleId = "TableGrid",
            Table = new DocumentTable { GridColumns = 1 },
        };
        var row = new DocumentTableRow();
        row.Cells.Add("cell");
        table.Table.Rows.Add(row);
        var tableSource = ExportDocx(table);
        AssertFailure(Add(tableSource, "cell", "cell", "updated"), "unsupported_document_tracked_replacement_target");
        AssertFailure(AddTableCell(tableSource, "cell", "cell", "updated", row: 2, column: 0), "document_tracked_replacement_target_not_found");
        AssertFailure(AddTableCell(tableSource, "cell", "cell", "updated", row: 0, column: 2), "document_tracked_replacement_target_not_found");
        AssertFailure(AddTableCell(tableSource, "cell", "cell", "updated", row: uint.MaxValue, column: 0), "document_tracked_replacement_target_not_found");
        AssertFailure(AddTableCell(tableSource, "cell", "cell", "updated", row: 0, column: uint.MaxValue), "document_tracked_replacement_target_not_found");

        var multipleParagraphs = EditDocx(tableSource, body =>
        {
            var cell = body.Elements<W.Table>().Single().Descendants<W.TableCell>().Single();
            cell.Append(new W.Paragraph(new W.Run(new W.Text("second paragraph"))));
        });
        AssertFailure(AddTableCell(multipleParagraphs, "cellsecond paragraph", "cell", "updated"), "unsupported_document_tracked_replacement_topology");

        var invalidMergeContinuation = EditDocx(tableSource, body =>
        {
            var cell = body.Elements<W.Table>().Single().Descendants<W.TableCell>().Single();
            var properties = cell.TableCellProperties ??= new W.TableCellProperties();
            properties.Append(new W.VerticalMerge { Val = W.MergedCellValues.Continue });
        });
        AssertFailure(AddTableCell(invalidMergeContinuation, "cell", "cell", "updated"), "unsupported_document_tracked_replacement_topology");

        var mismatchedTarget = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.AddDocxTrackedReplacement,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(tableSource),
            TrackedReplacement = new DocumentTrackedReplacementRequest
            {
                ExpectedSourceSha256 = Hash(tableSource),
                TargetBlockIndex = 0,
                Target = new DocumentTrackedReplacementTarget
                {
                    BlockIndex = 1,
                    TableCell = new DocumentTrackedReplacementTableCell { Row = 0, Column = 0 },
                },
                ExpectedParagraphText = "cell",
                Search = "cell",
                Replacement = "updated",
                Author = "Reviewer",
            },
        });
        AssertFailure(mismatchedTarget, "document_tracked_replacement_target_mismatch");
    }

    private static CodecResponse Add(
        byte[] source,
        string expectedText,
        string search,
        string replacement,
        string? expectedHash = null,
        uint targetBlockIndex = 0) => Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.AddDocxTrackedReplacement,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(source),
            TrackedReplacement = new DocumentTrackedReplacementRequest
            {
                ExpectedSourceSha256 = expectedHash ?? Hash(source),
                TargetBlockIndex = targetBlockIndex,
                ExpectedParagraphText = expectedText,
                Search = search,
                Replacement = replacement,
                Author = "Reviewer",
                Date = "2026-07-21T10:00:00Z",
            },
        });

    private static CodecResponse AddTableCell(
        byte[] source,
        string expectedText,
        string search,
        string replacement,
        uint row = 0,
        uint column = 0,
        uint blockIndex = 0) => Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.AddDocxTrackedReplacement,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(source),
            TrackedReplacement = new DocumentTrackedReplacementRequest
            {
                ExpectedSourceSha256 = Hash(source),
                TargetBlockIndex = blockIndex,
                Target = new DocumentTrackedReplacementTarget
                {
                    BlockIndex = blockIndex,
                    TableCell = new DocumentTrackedReplacementTableCell { Row = row, Column = column },
                },
                ExpectedParagraphText = expectedText,
                Search = search,
                Replacement = replacement,
                Author = "Reviewer",
                Date = "2026-07-21T10:00:00Z",
            },
        });

    private static CodecResponse Finalize(byte[] source, DocumentRevisionFinalizationMode mode)
    {
        var response = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.FinalizeDocxRevisions,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(source),
            RevisionFinalization = new DocumentRevisionFinalizationRequest
            {
                Mode = mode,
                ExpectedSourceSha256 = Hash(source),
            },
        });
        Assert.True(response.Ok, Diagnostics(response));
        return response;
    }

    private static byte[] ExportDocx(params DocumentBlock[] blocks)
    {
        var artifact = new ArtifactEnvelope
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Family = ArtifactFamily.Document,
            Document = new DocumentArtifact { Id = "document/1", Name = "Tracked replacement fixture" },
        };
        artifact.Document.Blocks.Add(blocks);
        var response = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = artifact,
        });
        Assert.True(response.Ok, Diagnostics(response));
        return response.File.ToByteArray();
    }

    private static byte[] EditDocx(byte[] source, Action<W.Body> edit)
    {
        using var stream = new MemoryStream();
        stream.Write(source);
        stream.Position = 0;
        using (var package = WordprocessingDocument.Open(stream, isEditable: true))
        {
            var document = package.MainDocumentPart!.Document!;
            edit(document.Body!);
            document.Save();
        }
        return stream.ToArray();
    }

    private static DocumentBlock Paragraph(string text, bool bold = false)
    {
        var block = new DocumentBlock
        {
            Id = "paragraph/1",
            StyleId = "Normal",
            Paragraph = new DocumentParagraph { Text = text },
        };
        block.Paragraph.Runs.Add(new DocumentRun { Text = text, Bold = bold });
        return block;
    }

    private static string BodyText(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        return string.Concat(package.MainDocumentPart!.Document!.Body!.Descendants<W.Text>().Select(value => value.Text));
    }

    private static string TableCellText(byte[] bytes, int row, int column)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var table = package.MainDocumentPart!.Document!.Body!.Elements<W.Table>().Single();
        var cell = table.Elements<W.TableRow>().ElementAt(row).Elements<W.TableCell>().ElementAt(column);
        return string.Concat(cell.Descendants<W.Text>().Select(value => value.Text));
    }

    private static void AssertNoRevisions(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        Assert.Empty(package.MainDocumentPart!.Document!.Descendants<W.InsertedRun>());
        Assert.Empty(package.MainDocumentPart.Document.Descendants<W.DeletedRun>());
    }

    private static void AssertFailure(CodecResponse response, string code)
    {
        Assert.False(response.Ok);
        Assert.Equal(code, Assert.Single(response.Diagnostics).Code);
        Assert.Empty(response.File);
    }

    private static CodecResponse Invoke(CodecRequest request) =>
        CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));

    private static string Diagnostics(CodecResponse response) =>
        string.Join("; ", response.Diagnostics.Select(item => $"{item.Code}: {item.Message}"));

    private static string HashText(string value) => Hash(System.Text.Encoding.UTF8.GetBytes(value));
    private static string Hash(byte[] bytes) => Hash(bytes.AsSpan());
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
