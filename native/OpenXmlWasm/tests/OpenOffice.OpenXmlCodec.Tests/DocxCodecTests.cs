using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OpenOffice.OpenXmlCodec.Tests;

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
                Assert.False(table.Source.Editable);
                Assert.Equal("Revenue", table.Table.Rows[0].Cells[0]);
                Assert.Equal("42", table.Table.Rows[0].Cells[1]);
            });
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
