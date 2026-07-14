using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OpenChestnut.Codec.Tests;

public sealed class DocxCodecTests
{
    [Fact]
    public void ProtocolRoundTripsMinimalDocument()
    {
        var exported = Invoke(ExportRequest());
        Assert.True(exported.Ok, Diagnostics(exported));
        Assert.Equal("PK", System.Text.Encoding.ASCII.GetString(exported.File.Span[..2]));
        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = WordprocessingDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));

        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = exported.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        Assert.Equal(ArtifactFamily.Document, imported.Artifact.Family);
        Assert.Collection(imported.Artifact.Document.Blocks,
            paragraph =>
            {
                Assert.Equal("Quarterly brief", paragraph.Paragraph.Text);
                Assert.True(paragraph.Source.Editable);
                Assert.Equal("Title", paragraph.StyleId);
                Assert.True(paragraph.Paragraph.Runs[0].Bold);
            },
            table =>
            {
                Assert.True(table.Source.Editable);
                Assert.NotEmpty(table.Source.ResidualSha256);
                Assert.Equal("Revenue", table.Table.Rows[0].Cells[0]);
                Assert.Equal("42", table.Table.Rows[0].Cells[1]);
            });
    }

    [Fact]
    public void SourcePreservingExportEditsFixedTopologyTableTextAndKeepsFormatting()
    {
        var authored = Invoke(ExportRequest());
        Assert.True(authored.Ok, Diagnostics(authored));
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(AddTableFormatting(authored.File.ToByteArray())),
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var block = imported.Artifact.Document.Blocks[1];
        Assert.Equal(DocumentBlock.ContentOneofCase.Table, block.ContentCase);
        Assert.True(block.Source.Editable);
        Assert.NotEmpty(block.Source.ResidualSha256);

        block.Table.Rows[0].Cells[0] = "Net revenue";
        block.Table.Rows[0].Cells[1] = "84";
        var exported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(exported.Ok, Diagnostics(exported));
        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var table = document.MainDocumentPart!.Document!.Body!.Elements<W.Table>().Single();
            var cells = table.Descendants<W.TableCell>().ToArray();
            Assert.Equal("Net revenue", cells[0].InnerText);
            Assert.Equal("84", cells[1].InnerText);
            Assert.Equal(W.JustificationValues.Center, cells[0].Descendants<W.Justification>().Single().Val?.Value);
            Assert.Equal("F2F4F7", cells[0].Descendants<W.Shading>().Single().Fill?.Value);
            Assert.NotNull(cells[0].Descendants<W.Italic>().SingleOrDefault());
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }
        var roundTrip = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = exported.File,
        });
        Assert.Equal("Net revenue", roundTrip.Artifact.Document.Blocks[1].Table.Rows[0].Cells[0]);
        Assert.Equal("84", roundTrip.Artifact.Document.Blocks[1].Table.Rows[0].Cells[1]);
    }

    [Fact]
    public void TableSliceRejectsComplexOrChangedTopologyAndSourceStyle()
    {
        var authored = Invoke(ExportRequest());
        var complex = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(AddSecondTableParagraph(authored.File.ToByteArray())),
        });
        Assert.True(complex.Ok, Diagnostics(complex));
        Assert.False(complex.Artifact.Document.Blocks[1].Source.Editable);
        var unchanged = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = complex.Artifact,
        });
        Assert.True(unchanged.Ok, Diagnostics(unchanged));
        complex.Artifact.Document.Blocks[1].Table.Rows[0].Cells[0] = "Unsafe complex edit";
        var complexRejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = complex.Artifact,
        });
        Assert.False(complexRejected.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(complexRejected.Diagnostics).Code);

        var topology = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        topology.Artifact.Document.Blocks[1].Table.Rows[0].Cells.RemoveAt(1);
        var topologyRejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = topology.Artifact,
        });
        Assert.False(topologyRejected.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(topologyRejected.Diagnostics).Code);

        var style = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        style.Artifact.Document.Blocks[1].StyleId = "LightShading";
        var styleRejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = style.Artifact,
        });
        Assert.False(styleRejected.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(styleRejected.Diagnostics).Code);
    }

    [Fact]
    public void SourcePreservingExportExposesAndEditsMergedTableOrigins()
    {
        var authored = Invoke(ExportRequest());
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(AddMergedTableGeometry(authored.File.ToByteArray())),
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var block = imported.Artifact.Document.Blocks[1];
        Assert.True(block.Source.Editable);
        Assert.Equal(3u, block.Table.GridColumns);
        Assert.Equal(3, block.Table.Rows.Count);

        var horizontal = block.Table.Rows[0].RichCells[0];
        Assert.Equal(0u, horizontal.GridColumn);
        Assert.Equal(2u, horizontal.ColumnSpan);
        Assert.Equal(1u, horizontal.RowSpan);
        Assert.Equal(DocumentTableVerticalMerge.None, horizontal.VerticalMerge);
        Assert.True(horizontal.Editable);

        var vertical = block.Table.Rows[0].RichCells[1];
        Assert.Equal(2u, vertical.GridColumn);
        Assert.Equal(1u, vertical.ColumnSpan);
        Assert.Equal(3u, vertical.RowSpan);
        Assert.Equal(DocumentTableVerticalMerge.Restart, vertical.VerticalMerge);
        Assert.True(vertical.Editable);
        var continuation = block.Table.Rows[1].RichCells[2];
        Assert.Equal(DocumentTableVerticalMerge.Continue, continuation.VerticalMerge);
        Assert.Equal(0u, continuation.RowSpan);
        Assert.False(continuation.Editable);

        block.Table.Rows[0].Cells[0] = "Edited horizontal origin";
        block.Table.Rows[0].Cells[1] = "Edited vertical origin";
        var exported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(exported.Ok, Diagnostics(exported));
        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var table = document.MainDocumentPart!.Document!.Body!.Elements<W.Table>().Single();
            var firstRow = table.Elements<W.TableRow>().First().Elements<W.TableCell>().ToArray();
            Assert.Equal("Edited horizontal origin", firstRow[0].InnerText);
            Assert.Equal(2, firstRow[0].TableCellProperties!.GridSpan!.Val!.Value);
            Assert.Equal("Edited vertical origin", firstRow[1].InnerText);
            Assert.Equal(W.MergedCellValues.Restart, firstRow[1].TableCellProperties!.VerticalMerge!.Val!.Value);
            Assert.Equal(2, table.Descendants<W.VerticalMerge>().Count(item => item.Val?.Value == W.MergedCellValues.Continue));
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }

        var roundTrip = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = exported.File,
        });
        Assert.Equal(3u, roundTrip.Artifact.Document.Blocks[1].Table.Rows[0].RichCells[1].RowSpan);

        block.Table.Rows[1].Cells[2] = "Unsafe continuation edit";
        var continuationRejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.False(continuationRejected.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(continuationRejected.Diagnostics).Code);

        block.Table.Rows[1].Cells[2] = string.Empty;
        block.Table.Rows[0].RichCells[0].ColumnSpan = 1;
        var topologyRejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.False(topologyRejected.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(topologyRejected.Diagnostics).Code);
    }

    [Fact]
    public void NonconformantVerticalMergeRemainsSourcePreservedAndReadOnly()
    {
        var authored = Invoke(ExportRequest());
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(AddMergedTableGeometry(authored.File.ToByteArray(), mismatchContinuation: true)),
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var block = imported.Artifact.Document.Blocks[1];
        Assert.False(block.Source.Editable);

        var unchanged = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(unchanged.Ok, Diagnostics(unchanged));

        block.Table.Rows[0].Cells[0] = "Unsafe malformed merge edit";
        var rejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void SourcePreservingExportEditsDirectNumberedParagraphTextAndKeepsDefinition()
    {
        var authored = Invoke(ExportRequest(includeSecondParagraph: true));
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(AddDirectNumbering(authored.File.ToByteArray())),
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var block = imported.Artifact.Document.Blocks[1];
        Assert.Equal(DocumentBlock.ContentOneofCase.Paragraph, block.ContentCase);
        Assert.True(block.Source.Editable);
        Assert.NotEmpty(block.Source.ResidualSha256);
        Assert.Equal(77u, block.Paragraph.Numbering.NumberingId);
        Assert.Equal(9u, block.Paragraph.Numbering.AbstractNumberingId);
        Assert.Equal(0u, block.Paragraph.Numbering.Level);
        Assert.Equal("upperLetter", block.Paragraph.Numbering.NumberFormat);
        Assert.Equal(3u, block.Paragraph.Numbering.Start);
        Assert.Equal("%1)", block.Paragraph.Numbering.LevelText);

        block.Paragraph.Text = "Edited numbered evidence";
        block.Paragraph.Runs[0].Text = block.Paragraph.Text;
        var exported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(exported.Ok, Diagnostics(exported));
        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var paragraph = document.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().ElementAt(1);
            Assert.Equal("Edited numbered evidence", paragraph.InnerText);
            Assert.Equal(77, paragraph.ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value);
            Assert.NotNull(paragraph.Descendants<W.Italic>().SingleOrDefault());
            var level = document.MainDocumentPart.NumberingDefinitionsPart!.Numbering!
                .Elements<W.AbstractNum>().Single().Elements<W.Level>().Single();
            Assert.Equal(W.NumberFormatValues.UpperLetter, level.NumberingFormat!.Val!.Value);
            Assert.Equal("%1)", level.LevelText!.Val!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }

        block.Paragraph.Numbering.Level = 1;
        var rejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void NumberedParagraphSlicePreservesButRejectsComplexTopologyEdits()
    {
        var authored = Invoke(ExportRequest(includeSecondParagraph: true));
        var source = AddDirectNumbering(authored.File.ToByteArray(), addSecondRun: true);
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(source),
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var block = imported.Artifact.Document.Blocks[1];
        Assert.NotNull(block.Paragraph.Numbering);
        Assert.False(block.Source.Editable);

        var unchanged = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(unchanged.Ok, Diagnostics(unchanged));
        block.Paragraph.Text = "Unsafe numbered edit";
        var rejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void SourcePreservingExportEditsStyleInheritedNumberedParagraphText()
    {
        var authored = Invoke(ExportRequest(includeSecondParagraph: true));
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(AddStyleInheritedNumbering(authored.File.ToByteArray())),
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var block = imported.Artifact.Document.Blocks[1];
        Assert.True(block.Source.Editable);
        Assert.Equal("DerivedList", block.StyleId);
        Assert.Equal(77u, block.Paragraph.Numbering.NumberingId);
        Assert.Equal(1u, block.Paragraph.Numbering.Level);
        Assert.Equal("lowerRoman", block.Paragraph.Numbering.NumberFormat);
        Assert.Equal(4u, block.Paragraph.Numbering.Start);
        Assert.Equal("%1.%2)", block.Paragraph.Numbering.LevelText);
        Assert.Equal(string.Empty, block.Paragraph.Numbering.NumberingStyleId);

        block.Paragraph.Text = "Edited inherited list evidence";
        block.Paragraph.Runs[0].Text = block.Paragraph.Text;
        var exported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(exported.Ok, Diagnostics(exported));
        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var paragraph = document.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().ElementAt(1);
            Assert.Equal("Edited inherited list evidence", paragraph.InnerText);
            Assert.Equal("DerivedList", paragraph.ParagraphProperties!.ParagraphStyleId!.Val!.Value);
            Assert.Null(paragraph.ParagraphProperties.NumberingProperties);
            Assert.Equal(77, document.MainDocumentPart.StyleDefinitionsPart!.Styles!
                .Elements<W.Style>().Single(style => style.StyleId == "BaseList")
                .StyleParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }

        var roundTrip = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = exported.File,
        });
        Assert.Equal("Edited inherited list evidence", roundTrip.Artifact.Document.Blocks[1].Paragraph.Text);
        Assert.Equal(1u, roundTrip.Artifact.Document.Blocks[1].Paragraph.Numbering.Level);
    }

    [Fact]
    public void SourcePreservingExportEditsNumberingStyleLinkedParagraphText()
    {
        var authored = Invoke(ExportRequest(includeSecondParagraph: true));
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(AddNumberingStyleLinkedNumbering(authored.File.ToByteArray())),
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var block = imported.Artifact.Document.Blocks[1];
        Assert.True(block.Source.Editable);
        Assert.Equal("DerivedList", block.StyleId);
        Assert.Equal(6u, block.Paragraph.Numbering.NumberingId);
        Assert.Equal(0u, block.Paragraph.Numbering.AbstractNumberingId);
        Assert.Equal(2u, block.Paragraph.Numbering.Level);
        Assert.Equal("upperRoman", block.Paragraph.Numbering.NumberFormat);
        Assert.Equal(9u, block.Paragraph.Numbering.Start);
        Assert.Equal("%1.%2.%3", block.Paragraph.Numbering.LevelText);
        Assert.Equal("AgentNumbering", block.Paragraph.Numbering.NumberingStyleId);

        block.Paragraph.Text = "Edited numbering-style link evidence";
        block.Paragraph.Runs[0].Text = block.Paragraph.Text;
        var exported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(exported.Ok, Diagnostics(exported));
        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var mainPart = document.MainDocumentPart!;
            Assert.Equal("Edited numbering-style link evidence", mainPart.Document!.Body!.Elements<W.Paragraph>().ElementAt(1).InnerText);
            Assert.Equal(4, mainPart.StyleDefinitionsPart!.Styles!.Elements<W.Style>()
                .Single(style => style.StyleId == "AgentNumbering")
                .StyleParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value);
            var abstracts = mainPart.NumberingDefinitionsPart!.Numbering!.Elements<W.AbstractNum>().ToArray();
            Assert.Equal("AgentNumbering", abstracts.Single(item => item.AbstractNumberId?.Value == 0).NumberingStyleLink!.Val!.Value);
            Assert.Equal("AgentNumbering", abstracts.Single(item => item.AbstractNumberId?.Value == 2).StyleLink!.Val!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }

        var roundTrip = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = exported.File,
        });
        Assert.Equal("AgentNumbering", roundTrip.Artifact.Document.Blocks[1].Paragraph.Numbering.NumberingStyleId);
        Assert.Equal(2u, roundTrip.Artifact.Document.Blocks[1].Paragraph.Numbering.Level);

        block.Paragraph.Numbering.NumberingStyleId = "UnsafeReplacement";
        var rejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void CyclicNumberingStyleLinkRemainsSourcePreservedAndReadOnly()
    {
        var authored = Invoke(ExportRequest(includeSecondParagraph: true));
        var source = AddNumberingStyleLinkedNumbering(authored.File.ToByteArray(), cyclic: true);
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(source),
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var block = imported.Artifact.Document.Blocks[1];
        Assert.False(block.Source.Editable);
        Assert.Equal(6u, block.Paragraph.Numbering.NumberingId);

        var unchanged = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(unchanged.Ok, Diagnostics(unchanged));
        Assert.Equal(source, unchanged.File.ToByteArray());

        block.Paragraph.Text = "Unsafe numbering-style cycle edit";
        block.Paragraph.Runs[0].Text = block.Paragraph.Text;
        var rejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void CyclicStyleInheritedNumberingRemainsSourcePreservedAndReadOnly()
    {
        var authored = Invoke(ExportRequest(includeSecondParagraph: true));
        var source = AddStyleNumberingCycle(AddStyleInheritedNumbering(authored.File.ToByteArray()));
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(source),
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var block = imported.Artifact.Document.Blocks[1];
        Assert.Equal(DocumentBlock.ContentOneofCase.Paragraph, block.ContentCase);
        Assert.False(block.Source.Editable);
        Assert.Null(block.Paragraph.Numbering);

        var unchanged = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(unchanged.Ok, Diagnostics(unchanged));
        Assert.Equal(source, unchanged.File.ToByteArray());

        block.Paragraph.Text = "Unsafe cyclic list edit";
        block.Paragraph.Runs[0].Text = block.Paragraph.Text;
        var rejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void SourcePreservingExportKeepsAdvancedParagraphAndRelationships()
    {
        var first = Invoke(ExportRequest(includeSecondParagraph: true));
        var source = AddBookmarkAndExternalRelationship(first.File.ToByteArray());
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(source),
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        Assert.False(imported.Artifact.Document.Blocks[0].Source.Editable);
        Assert.True(imported.Artifact.Document.Blocks[1].Source.Editable);
        Assert.Contains(imported.Artifact.OpaqueOpc.PackageRelationships,
            relationship => relationship.SourcePath == "word/document.xml" && relationship.Id == "rIdExternal");

        var editable = imported.Artifact.Document.Blocks[1].Paragraph;
        editable.Text = "Edited safely";
        editable.Runs[0].Text = editable.Text;
        var preserved = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(preserved.Ok, Diagnostics(preserved));
        Assert.Equal("opaque_content_preserved", Assert.Single(preserved.Diagnostics).Code);

        using (var stream = new MemoryStream(preserved.File.ToByteArray()))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var mainPart = document.MainDocumentPart!;
            var body = mainPart.Document!.Body!;
            Assert.NotNull(body.Descendants<W.BookmarkStart>().SingleOrDefault(item => item.Name?.Value == "preserve-me"));
            Assert.Equal("Edited safely", body.Elements<W.Paragraph>().ElementAt(1).InnerText);
            Assert.Contains(mainPart.ExternalRelationships,
                relationship => relationship.Id == "rIdExternal" && relationship.Uri == new Uri("https://example.invalid/data"));
        }

        imported.Artifact.Document.Blocks[0].Paragraph.Text = "Unsafe edit";
        var rejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void SourcePreservingExportRejectsTamperedBinding()
    {
        var exported = Invoke(ExportRequest());
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = exported.File,
        });
        imported.Artifact.Document.Blocks[0].Source.SemanticSha256 = new string('0', 64);
        var response = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.False(response.Ok);
        Assert.Equal("document_source_semantics_mismatch", Assert.Single(response.Diagnostics).Code);
    }

    [Fact]
    public void SourcePreservingExportRejectsInvalidOwnedMarkup()
    {
        var exported = Invoke(ExportRequest());
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(AddInvalidParagraphMarkup(exported.File.ToByteArray())),
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var response = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.False(response.Ok);
        Assert.Equal("openxml_validation_failed", Assert.Single(response.Diagnostics).Code);
    }

    [Fact]
    public void SourcePreservingExportEditsOwnedHyperlinkTextTargetsAndRelationships()
    {
        var authored = Invoke(HyperlinkExportRequest());
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = AddBookmarkTarget(authored.File.ToByteArray());
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(source),
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var block = imported.Artifact.Document.Blocks[0];
        Assert.Equal(DocumentBlock.ContentOneofCase.Hyperlink, block.ContentCase);
        Assert.True(block.Source.Editable);
        Assert.NotEmpty(block.Source.ResidualSha256);
        Assert.Equal("https://example.test/original", block.Hyperlink.ExternalUri);
        var originalRelationshipId = block.Hyperlink.RelationshipId;

        block.Hyperlink.Text = "Updated external link";
        block.Hyperlink.ExternalUri = "https://example.test/updated";
        block.Hyperlink.Tooltip = "Updated target";
        block.Hyperlink.History = false;
        var external = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(external.Ok, Diagnostics(external));
        using (var stream = new MemoryStream(external.File.ToByteArray()))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var mainPart = document.MainDocumentPart!;
            var hyperlink = mainPart.Document!.Body!.Elements<W.Paragraph>().First().Elements<W.Hyperlink>().Single();
            Assert.Equal("Updated external link", hyperlink.InnerText);
            Assert.Equal("Updated target", hyperlink.Tooltip?.Value);
            Assert.False(hyperlink.History?.Value);
            Assert.NotEqual(originalRelationshipId, hyperlink.Id?.Value);
            Assert.Equal("https://example.test/updated", mainPart.HyperlinkRelationships.Single(item => item.Id == hyperlink.Id).Uri.OriginalString);
            Assert.DoesNotContain(mainPart.HyperlinkRelationships, item => item.Id == originalRelationshipId);
            Assert.Equal("0000FF", hyperlink.Descendants<W.Color>().Single().Val?.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }

        var internalImport = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(source),
        });
        var internalBlock = internalImport.Artifact.Document.Blocks[0];
        internalBlock.Hyperlink.Text = "Jump to target";
        internalBlock.Hyperlink.InternalAnchor = "TargetBookmark";
        internalBlock.Hyperlink.Tooltip = "Internal target";
        var internalExport = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = internalImport.Artifact,
        });
        Assert.True(internalExport.Ok, Diagnostics(internalExport));
        using (var stream = new MemoryStream(internalExport.File.ToByteArray()))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var mainPart = document.MainDocumentPart!;
            var hyperlink = mainPart.Document!.Body!.Elements<W.Paragraph>().First().Elements<W.Hyperlink>().Single();
            Assert.Equal("TargetBookmark", hyperlink.Anchor?.Value);
            Assert.Null(hyperlink.Id);
            Assert.Empty(mainPart.HyperlinkRelationships);
            Assert.NotNull(mainPart.Document.Descendants<W.BookmarkStart>().SingleOrDefault(item => item.Name?.Value == "TargetBookmark"));
        }

        var rebound = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = internalExport.File,
        });
        rebound.Artifact.Document.Blocks[0].Hyperlink.ExternalUri = "https://example.test/rebound";
        var reboundExport = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = rebound.Artifact,
        });
        Assert.True(reboundExport.Ok, Diagnostics(reboundExport));
        using (var stream = new MemoryStream(reboundExport.File.ToByteArray()))
        using (var document = WordprocessingDocument.Open(stream, false))
            Assert.Equal("https://example.test/rebound", Assert.Single(document.MainDocumentPart!.HyperlinkRelationships).Uri.OriginalString);
    }

    [Fact]
    public void HyperlinkSliceRejectsUnsupportedTopologyAndUnsafeUri()
    {
        var authored = Invoke(HyperlinkExportRequest());
        var complex = AddSecondHyperlinkRun(authored.File.ToByteArray());
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(complex),
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        Assert.Equal(DocumentBlock.ContentOneofCase.Hyperlink, imported.Artifact.Document.Blocks[0].ContentCase);
        Assert.False(imported.Artifact.Document.Blocks[0].Source.Editable);
        imported.Artifact.Document.Blocks[0].Hyperlink.Text = "Unsafe topology edit";
        var rejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(rejected.Diagnostics).Code);

        var locatorImport = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        locatorImport.Artifact.Document.Blocks[0].Hyperlink.RelationshipId = string.Empty;
        var locatorRejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = locatorImport.Artifact,
        });
        Assert.False(locatorRejected.Ok);
        Assert.Equal("document_source_binding_mismatch", Assert.Single(locatorRejected.Diagnostics).Code);

        var unsafeRequest = HyperlinkExportRequest();
        unsafeRequest.Artifact.Document.Blocks[0].Hyperlink.ExternalUri = "javascript:alert(1)";
        var unsafeResponse = Invoke(unsafeRequest);
        Assert.False(unsafeResponse.Ok);
        Assert.Equal("invalid_document_hyperlink", Assert.Single(unsafeResponse.Diagnostics).Code);
    }

    [Fact]
    public void SourceHyperlinkEditRetainsSharedRelationshipForOtherOwners()
    {
        var request = HyperlinkExportRequest();
        request.Artifact.Document.Blocks.Insert(1, new DocumentBlock
        {
            Id = "document/shared-link",
            StyleId = "Normal",
            Hyperlink = new DocumentHyperlink
            {
                Text = "Shared original target",
                ExternalUri = "https://example.test/original",
            },
        });
        var authored = Invoke(request);
        Assert.True(authored.Ok, Diagnostics(authored));
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        var originalRelationshipId = imported.Artifact.Document.Blocks[0].Hyperlink.RelationshipId;
        Assert.Equal(originalRelationshipId, imported.Artifact.Document.Blocks[1].Hyperlink.RelationshipId);
        imported.Artifact.Document.Blocks[0].Hyperlink.ExternalUri = "https://example.test/first-only";
        var exported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(exported.Ok, Diagnostics(exported));
        using var stream = new MemoryStream(exported.File.ToByteArray());
        using var document = WordprocessingDocument.Open(stream, false);
        var mainPart = document.MainDocumentPart!;
        var hyperlinks = mainPart.Document!.Descendants<W.Hyperlink>().ToArray();
        Assert.Equal(2, mainPart.HyperlinkRelationships.Count());
        Assert.Equal(originalRelationshipId, hyperlinks[1].Id?.Value);
        Assert.Equal("https://example.test/original", mainPart.HyperlinkRelationships.Single(item => item.Id == originalRelationshipId).Uri.OriginalString);
        Assert.Equal("https://example.test/first-only", mainPart.HyperlinkRelationships.Single(item => item.Id == hyperlinks[0].Id).Uri.OriginalString);
    }

    [Fact]
    public void SourcePreservingExportEditsBoundedSimpleFieldAndKeepsFormatting()
    {
        var authored = Invoke(FieldExportRequest());
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = AddSimpleFieldFormatting(authored.File.ToByteArray());
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(source),
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var block = imported.Artifact.Document.Blocks[0];
        Assert.Equal(DocumentBlock.ContentOneofCase.Field, block.ContentCase);
        Assert.True(block.Source.Editable);
        Assert.NotEmpty(block.Source.ResidualSha256);
        Assert.Equal("PAGE", block.Field.Instruction);
        Assert.Equal("1", block.Field.Display);

        block.Field.Instruction = "NUMPAGES \\* MERGEFORMAT";
        block.Field.Display = "2";
        var exported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(exported.Ok, Diagnostics(exported));
        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var field = document.MainDocumentPart!.Document!.Descendants<W.SimpleField>().Single();
            Assert.Equal("NUMPAGES \\* MERGEFORMAT", field.Instruction?.Value);
            Assert.Equal("2", field.InnerText);
            Assert.True(field.Dirty?.Value);
            Assert.NotNull(field.Descendants<W.Bold>().SingleOrDefault());
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }
        var roundTrip = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = exported.File,
        });
        Assert.Equal("NUMPAGES \\* MERGEFORMAT", roundTrip.Artifact.Document.Blocks[0].Field.Instruction);
        Assert.Equal("2", roundTrip.Artifact.Document.Blocks[0].Field.Display);
    }

    [Fact]
    public void FieldSlicePreservesUnsupportedInstructionsAndTopologiesReadOnly()
    {
        var authored = Invoke(FieldExportRequest());
        var unsupportedSource = SetSimpleFieldInstruction(authored.File.ToByteArray(), "INCLUDETEXT \\\"https://example.test/data\\\"");
        var unsupported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(unsupportedSource),
        });
        Assert.True(unsupported.Ok, Diagnostics(unsupported));
        Assert.Equal(DocumentBlock.ContentOneofCase.Field, unsupported.Artifact.Document.Blocks[0].ContentCase);
        Assert.False(unsupported.Artifact.Document.Blocks[0].Source.Editable);
        var unchanged = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = unsupported.Artifact,
        });
        Assert.True(unchanged.Ok, Diagnostics(unchanged));
        unsupported.Artifact.Document.Blocks[0].Field.Display = "Unsafe edit";
        var rejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = unsupported.Artifact,
        });
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(rejected.Diagnostics).Code);

        var multiRun = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(AddSecondFieldRun(authored.File.ToByteArray())),
        });
        Assert.Equal(DocumentBlock.ContentOneofCase.Field, multiRun.Artifact.Document.Blocks[0].ContentCase);
        Assert.False(multiRun.Artifact.Document.Blocks[0].Source.Editable);

        var unsafeRequest = FieldExportRequest();
        unsafeRequest.Artifact.Document.Blocks[0].Field.Instruction = "DDEAUTO command";
        var unsafeResponse = Invoke(unsafeRequest);
        Assert.False(unsafeResponse.Ok);
        Assert.Equal("invalid_document_field", Assert.Single(unsafeResponse.Diagnostics).Code);
    }

    private static CodecResponse Invoke(CodecRequest request) =>
        CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));

    private static string Diagnostics(CodecResponse response) =>
        string.Join("\n", response.Diagnostics.Select(item => $"{item.Code}: {item.Message}"));

    private static CodecRequest ExportRequest(bool includeSecondParagraph = false)
    {
        var title = new DocumentBlock
        {
            Id = "document/title",
            StyleId = "Title",
            Paragraph = new DocumentParagraph { Text = "Quarterly brief" },
        };
        title.Paragraph.Runs.Add(new DocumentRun { Text = title.Paragraph.Text, Bold = true });
        var document = new DocumentArtifact { Id = "document/test", Name = "Quarterly brief" };
        document.Blocks.Add(title);
        if (includeSecondParagraph)
        {
            var second = new DocumentBlock
            {
                Id = "document/second",
                Paragraph = new DocumentParagraph { Text = "Editable paragraph" },
            };
            second.Paragraph.Runs.Add(new DocumentRun { Text = second.Paragraph.Text });
            document.Blocks.Add(second);
        }
        else
        {
            var table = new DocumentBlock { Id = "document/table", StyleId = "TableGrid", Table = new DocumentTable() };
            var row = new DocumentTableRow();
            row.Cells.Add("Revenue");
            row.Cells.Add("42");
            table.Table.Rows.Add(row);
            document.Blocks.Add(table);
        }
        return new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = new ArtifactEnvelope
            {
                ProtocolVersion = CodecProtocol.ProtocolVersion,
                Family = ArtifactFamily.Document,
                Document = document,
            },
        };
    }

    private static CodecRequest HyperlinkExportRequest()
    {
        var document = new DocumentArtifact { Id = "document/hyperlink", Name = "Hyperlink fixture" };
        document.Blocks.Add(new DocumentBlock
        {
            Id = "document/link",
            StyleId = "Normal",
            Hyperlink = new DocumentHyperlink
            {
                Text = "Original link",
                ExternalUri = "https://example.test/original",
            },
        });
        var target = new DocumentBlock
        {
            Id = "document/target",
            Paragraph = new DocumentParagraph { Text = "Target paragraph" },
        };
        target.Paragraph.Runs.Add(new DocumentRun { Text = target.Paragraph.Text });
        document.Blocks.Add(target);
        return new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = new ArtifactEnvelope
            {
                ProtocolVersion = CodecProtocol.ProtocolVersion,
                Family = ArtifactFamily.Document,
                Document = document,
            },
        };
    }

    private static CodecRequest FieldExportRequest()
    {
        var document = new DocumentArtifact { Id = "document/field", Name = "Field fixture" };
        document.Blocks.Add(new DocumentBlock
        {
            Id = "document/page-field",
            StyleId = "Normal",
            Field = new DocumentField { Instruction = "PAGE", Display = "1" },
        });
        return new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = new ArtifactEnvelope
            {
                ProtocolVersion = CodecProtocol.ProtocolVersion,
                Family = ArtifactFamily.Document,
                Document = document,
            },
        };
    }

    private static byte[] AddSimpleFieldFormatting(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var field = document.MainDocumentPart!.Document!.Descendants<W.SimpleField>().Single();
            field.Dirty = true;
            field.Elements<W.Run>().Single().PrependChild(new W.RunProperties(new W.Bold()));
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddTableFormatting(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var cell = document.MainDocumentPart!.Document!.Body!.Elements<W.Table>().Single().Descendants<W.TableCell>().First();
            cell.PrependChild(new W.TableCellProperties(new W.Shading { Val = W.ShadingPatternValues.Clear, Fill = "F2F4F7" }));
            var paragraph = cell.Elements<W.Paragraph>().Single();
            paragraph.PrependChild(new W.ParagraphProperties(new W.Justification { Val = W.JustificationValues.Center }));
            paragraph.Elements<W.Run>().Single().PrependChild(new W.RunProperties(new W.Italic()));
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddSecondTableParagraph(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var cell = document.MainDocumentPart!.Document!.Body!.Elements<W.Table>().Single().Descendants<W.TableCell>().First();
            cell.Append(new W.Paragraph(new W.Run(new W.Text(" detail"))));
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddMergedTableGeometry(byte[] bytes, bool mismatchContinuation = false)
    {
        static W.TableCell Cell(string text, int span = 1, W.MergedCellValues? merge = null)
        {
            var properties = new W.TableCellProperties();
            if (span > 1) properties.Append(new W.GridSpan { Val = span });
            if (merge is not null) properties.Append(new W.VerticalMerge { Val = merge });
            return new W.TableCell(properties, new W.Paragraph(new W.Run(new W.Text(text))));
        }

        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var table = document.MainDocumentPart!.Document!.Body!.Elements<W.Table>().Single();
            var replacement = new W.Table();
            if (table.TableProperties is not null)
                replacement.Append(table.TableProperties.CloneNode(true));
            replacement.Append(new W.TableGrid(new W.GridColumn(), new W.GridColumn(), new W.GridColumn()));
            replacement.Append(new W.TableRow(
                Cell("Horizontal origin", span: 2),
                Cell("Vertical origin", merge: W.MergedCellValues.Restart)));
            replacement.Append(mismatchContinuation
                ? new W.TableRow(Cell("A"), Cell(string.Empty, span: 2, merge: W.MergedCellValues.Continue))
                : new W.TableRow(Cell("A"), Cell("B"), Cell(string.Empty, merge: W.MergedCellValues.Continue)));
            replacement.Append(new W.TableRow(
                Cell("Footer", span: 2),
                Cell(string.Empty, merge: W.MergedCellValues.Continue)));
            table.InsertBeforeSelf(replacement);
            table.Remove();
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddDirectNumbering(byte[] bytes, bool addSecondRun = false)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var mainPart = document.MainDocumentPart!;
            var paragraph = mainPart.Document!.Body!.Elements<W.Paragraph>().ElementAt(1);
            paragraph.ParagraphProperties = new W.ParagraphProperties(
                new W.NumberingProperties(
                    new W.NumberingLevelReference { Val = 0 },
                    new W.NumberingId { Val = 77 }));
            paragraph.Elements<W.Run>().Single().PrependChild(new W.RunProperties(new W.Italic()));
            if (addSecondRun) paragraph.Append(new W.Run(new W.Text(" detail")));

            var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new W.Numbering(
                new W.AbstractNum(
                    new W.Level(
                        new W.StartNumberingValue { Val = 3 },
                        new W.NumberingFormat { Val = W.NumberFormatValues.UpperLetter },
                        new W.LevelText { Val = "%1)" }) { LevelIndex = 0 }) { AbstractNumberId = 9 },
                new W.NumberingInstance(new W.AbstractNumId { Val = 9 }) { NumberID = 77 });
            mainPart.Document.Save();
            numberingPart.Numbering.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddStyleInheritedNumbering(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var mainPart = document.MainDocumentPart!;
            var paragraph = mainPart.Document!.Body!.Elements<W.Paragraph>().ElementAt(1);
            paragraph.ParagraphProperties = new W.ParagraphProperties(
                new W.ParagraphStyleId { Val = "DerivedList" });

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new W.Styles(
                new W.Style(
                    new W.StyleName { Val = "Base list" },
                    new W.StyleParagraphProperties(
                        new W.NumberingProperties(new W.NumberingId { Val = 77 })))
                { Type = W.StyleValues.Paragraph, StyleId = "BaseList" },
                new W.Style(
                    new W.StyleName { Val = "Derived list" },
                    new W.BasedOn { Val = "BaseList" },
                    new W.StyleParagraphProperties(
                        new W.NumberingProperties(new W.NumberingLevelReference { Val = 8 })))
                { Type = W.StyleValues.Paragraph, StyleId = "DerivedList" });

            var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new W.Numbering(
                new W.AbstractNum(
                    new W.Level(
                        new W.StartNumberingValue { Val = 1 },
                        new W.NumberingFormat { Val = W.NumberFormatValues.UpperLetter },
                        new W.ParagraphStyleIdInLevel { Val = "BaseList" },
                        new W.LevelText { Val = "%1)" }) { LevelIndex = 0 },
                    new W.Level(
                        new W.StartNumberingValue { Val = 4 },
                        new W.NumberingFormat { Val = W.NumberFormatValues.LowerRoman },
                        new W.ParagraphStyleIdInLevel { Val = "DerivedList" },
                        new W.LevelText { Val = "%1.%2)" }) { LevelIndex = 1 }) { AbstractNumberId = 9 },
                new W.NumberingInstance(new W.AbstractNumId { Val = 9 }) { NumberID = 77 });
            mainPart.Document.Save();
            stylesPart.Styles.Save();
            numberingPart.Numbering.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddNumberingStyleLinkedNumbering(byte[] bytes, bool cyclic = false)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var mainPart = document.MainDocumentPart!;
            var paragraph = mainPart.Document!.Body!.Elements<W.Paragraph>().ElementAt(1);
            paragraph.ParagraphProperties = new W.ParagraphProperties(
                new W.ParagraphStyleId { Val = "DerivedList" });

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new W.Styles(
                new W.Style(
                    new W.StyleName { Val = "Base list" },
                    new W.StyleParagraphProperties(
                        new W.NumberingProperties(new W.NumberingId { Val = 6 })))
                { Type = W.StyleValues.Paragraph, StyleId = "BaseList" },
                new W.Style(
                    new W.StyleName { Val = "Derived list" },
                    new W.BasedOn { Val = "BaseList" },
                    new W.StyleParagraphProperties(
                        new W.NumberingProperties(new W.NumberingLevelReference { Val = 8 })))
                { Type = W.StyleValues.Paragraph, StyleId = "DerivedList" },
                new W.Style(
                    new W.StyleName { Val = "Agent numbering" },
                    new W.StyleParagraphProperties(
                        new W.NumberingProperties(new W.NumberingId { Val = cyclic ? 6 : 4 })))
                { Type = W.StyleValues.Numbering, StyleId = "AgentNumbering" });

            var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new W.Numbering(
                new W.AbstractNum(
                    new W.MultiLevelType { Val = W.MultiLevelValues.Multilevel },
                    new W.NumberingStyleLink { Val = "AgentNumbering" })
                { AbstractNumberId = 0 },
                new W.AbstractNum(
                    new W.MultiLevelType { Val = W.MultiLevelValues.Multilevel },
                    new W.StyleLink { Val = "AgentNumbering" },
                    new W.Level(
                        new W.StartNumberingValue { Val = 1 },
                        new W.NumberingFormat { Val = W.NumberFormatValues.Decimal },
                        new W.ParagraphStyleIdInLevel { Val = "BaseList" },
                        new W.LevelText { Val = "%1." }) { LevelIndex = 0 },
                    new W.Level(
                        new W.StartNumberingValue { Val = 3 },
                        new W.NumberingFormat { Val = W.NumberFormatValues.UpperRoman },
                        new W.ParagraphStyleIdInLevel { Val = "DerivedList" },
                        new W.LevelText { Val = "%1.%2.%3" }) { LevelIndex = 2 })
                { AbstractNumberId = 2 },
                new W.NumberingInstance(new W.AbstractNumId { Val = 2 }) { NumberID = 4 },
                new W.NumberingInstance(
                    new W.AbstractNumId { Val = 0 },
                    new W.LevelOverride(
                        new W.StartOverrideNumberingValue { Val = 9 }) { LevelIndex = 2 })
                { NumberID = 6 });
            mainPart.Document.Save();
            stylesPart.Styles.Save();
            numberingPart.Numbering.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddStyleNumberingCycle(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var styles = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            var baseStyle = styles.Elements<W.Style>().Single(style => style.StyleId == "BaseList");
            baseStyle.InsertAfter(new W.BasedOn { Val = "DerivedList" }, baseStyle.StyleName);
            styles.Save();
        }
        return stream.ToArray();
    }

    private static byte[] SetSimpleFieldInstruction(byte[] bytes, string instruction)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            document.MainDocumentPart!.Document!.Descendants<W.SimpleField>().Single().Instruction = instruction;
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddSecondFieldRun(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            document.MainDocumentPart!.Document!.Descendants<W.SimpleField>().Single().Append(new W.Run(new W.Text(" second result")));
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddBookmarkTarget(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var paragraph = document.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().ElementAt(1);
            paragraph.PrependChild(new W.BookmarkStart { Id = "41", Name = "TargetBookmark" });
            paragraph.Append(new W.BookmarkEnd { Id = "41" });
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddSecondHyperlinkRun(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var hyperlink = document.MainDocumentPart!.Document!.Descendants<W.Hyperlink>().Single();
            hyperlink.Append(new W.Run(new W.Text(" second run")));
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddBookmarkAndExternalRelationship(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var mainPart = document.MainDocumentPart!;
            var paragraph = mainPart.Document!.Body!.Elements<W.Paragraph>().First();
            var bookmark = new W.BookmarkStart { Id = "1", Name = "preserve-me" };
            if (paragraph.ParagraphProperties is { } properties) properties.InsertAfterSelf(bookmark);
            else paragraph.PrependChild(bookmark);
            paragraph.Append(new W.BookmarkEnd { Id = "1" });
            mainPart.AddExternalRelationship(
                "urn:open-office-artifact-tool:test",
                new Uri("https://example.invalid/data"),
                "rIdExternal");
            mainPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddInvalidParagraphMarkup(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("word/document.xml") ?? throw new InvalidOperationException("Document part is missing.");
            string xml;
            using (var reader = new StreamReader(entry.Open())) xml = reader.ReadToEnd();
            var closing = xml.IndexOf("</w:p>", StringComparison.Ordinal);
            entry.Delete();
            var replacement = archive.CreateEntry("word/document.xml");
            using var writer = new StreamWriter(replacement.Open());
            writer.Write(xml.Insert(closing, "<w:notARealParagraphChild/>"));
        }
        return stream.ToArray();
    }
}
