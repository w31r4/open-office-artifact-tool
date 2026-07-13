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
