using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using B = DocumentFormat.OpenXml.Bibliography;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using W16Cid = DocumentFormat.OpenXml.Office2019.Word.Cid;
using W16Cex = DocumentFormat.OpenXml.Office2021.Word.CommentsExt;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OpenChestnut.Codec.Tests;

public sealed class DocxCodecTests
{
    [Fact]
    public void ImportedBlockBindingRequiresValidatedSourcePackageSnapshot()
    {
        var authored = Invoke(OfficeSkillProfileExportRequest());
        Assert.True(authored.Ok, Diagnostics(authored));
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        Assert.Contains(imported.Artifact.Document.Blocks, block => block.Source is not null);

        imported.Artifact.Source = null;
        imported.Artifact.OpaqueOpc.SourcePackage = new SourcePackageSnapshot();
        var rejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });

        Assert.False(rejected.Ok);
        Assert.Equal("missing_source_package", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void SourceFreeParagraphReimportsTwoIndependentlyFormattedRuns()
    {
        var document = new DocumentArtifact { Id = "document/two-runs", Name = "Two formatted runs" };
        var paragraph = new DocumentBlock
        {
            Id = "document/two-runs/paragraph",
            StyleId = "Normal",
            Paragraph = new DocumentParagraph { Text = "Bold red" },
        };
        paragraph.Paragraph.Runs.Add(new DocumentRun
        {
            Text = "Bold ",
            Formatting = new DocumentRunFormatting { Bold = true },
        });
        paragraph.Paragraph.Runs.Add(new DocumentRun
        {
            Text = "red",
            Formatting = new DocumentRunFormatting { Italic = true, ColorRgb = "CC0000" },
        });
        document.Blocks.Add(paragraph);
        var authored = Invoke(new CodecRequest
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
        });
        Assert.True(authored.Ok, Diagnostics(authored));

        var imported = DocxCodec.Import(authored.File.ToByteArray(), EffectiveCodecLimits.From(null)).Artifact;
        var runs = Assert.Single(imported.Document.Blocks).Paragraph.Runs;
        Assert.Collection(runs,
            run =>
            {
                Assert.Equal("Bold ", run.Text);
                Assert.True(run.Formatting.Bold);
            },
            run =>
            {
                Assert.Equal("red", run.Text);
                Assert.True(run.Formatting.Italic);
                Assert.Equal("CC0000", run.Formatting.ColorRgb);
            });
    }

    [Fact]
    public void InlineSeqAndRefFieldsAuthorImportEditAndProtectSourceTopology()
    {
        var document = new DocumentArtifact { Id = "document/inline-fields", Name = "Inline fields" };
        var paragraph = new DocumentBlock
        {
            Id = "document/inline-fields/caption",
            StyleId = "Caption",
            Paragraph = new DocumentParagraph { Text = "Figure 0: Revenue. See 0." },
        };
        paragraph.Paragraph.Runs.Add(new DocumentRun { Text = "Figure " });
        paragraph.Paragraph.Runs.Add(new DocumentRun
        {
            Text = "0",
            InlineField = new DocumentInlineField
            {
                Instruction = "SEQ Figure \\* ARABIC",
                BookmarkName = "fig1",
            },
            Formatting = new DocumentRunFormatting { Bold = true },
        });
        paragraph.Paragraph.Runs.Add(new DocumentRun { Text = ": Revenue. See " });
        paragraph.Paragraph.Runs.Add(new DocumentRun
        {
            Text = "0",
            InlineField = new DocumentInlineField { Instruction = "REF fig1 \\h" },
        });
        paragraph.Paragraph.Runs.Add(new DocumentRun { Text = "." });
        document.Blocks.Add(paragraph);

        var authored = Invoke(new CodecRequest
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
        });
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var word = WordprocessingDocument.Open(stream, false))
        {
            var nativeParagraph = word.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().Single();
            Assert.Equal(13, nativeParagraph.Elements<W.Run>().Count());
            Assert.Equal(
                [" SEQ Figure \\* ARABIC ", " REF fig1 \\h "],
                nativeParagraph.Descendants<W.FieldCode>().Select(code => code.Text).ToArray());
            var children = nativeParagraph.ChildElements.ToArray();
            var bookmarkStartIndex = Array.FindIndex(children, child => child is W.BookmarkStart);
            var bookmarkEndIndex = Array.FindIndex(children, child => child is W.BookmarkEnd);
            Assert.Equal("fig1", Assert.IsType<W.BookmarkStart>(children[bookmarkStartIndex]).Name!.Value);
            Assert.Equal("0", Assert.IsType<W.BookmarkStart>(children[bookmarkStartIndex]).Id!.Value);
            Assert.Equal(bookmarkStartIndex + 2, bookmarkEndIndex);
            Assert.IsType<W.Run>(children[bookmarkStartIndex + 1]);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(word));
        }

        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var importedBlock = Assert.Single(imported.Artifact.Document.Blocks);
        Assert.True(importedBlock.Source.Editable);
        Assert.Equal(5, importedBlock.Paragraph.Runs.Count);
        Assert.Equal("SEQ Figure \\* ARABIC", importedBlock.Paragraph.Runs[1].InlineField.Instruction);
        Assert.Equal("fig1", importedBlock.Paragraph.Runs[1].InlineField.BookmarkName);
        Assert.Equal("0", importedBlock.Paragraph.Runs[1].InlineField.BookmarkNativeId);
        Assert.Equal("REF fig1 \\h", importedBlock.Paragraph.Runs[3].InlineField.Instruction);

        importedBlock.Paragraph.Runs[1].Text = "1";
        importedBlock.Paragraph.Runs[2].Text = ": Updated revenue. See ";
        importedBlock.Paragraph.Runs[3].Text = "1";
        importedBlock.Paragraph.Text = string.Concat(importedBlock.Paragraph.Runs.Select(run => run.Text));
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        var secondImport = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = edited.File,
        });
        Assert.Equal("Figure 1: Updated revenue. See 1.", secondImport.Artifact.Document.Blocks[0].Paragraph.Text);
        Assert.Equal("fig1", secondImport.Artifact.Document.Blocks[0].Paragraph.Runs[1].InlineField.BookmarkName);

        importedBlock.Paragraph.Runs[1].InlineField.BookmarkName = "fig2";
        var bookmarkTopologyRejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.False(bookmarkTopologyRejected.Ok);
        Assert.Equal("document_inline_field_topology_changed", Assert.Single(bookmarkTopologyRejected.Diagnostics).Code);
        importedBlock.Paragraph.Runs[1].InlineField.BookmarkName = "fig1";
        importedBlock.Paragraph.Runs[3].InlineField.Instruction = "REF fig2 \\h";
        var topologyRejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.False(topologyRejected.Ok);
        Assert.Equal("document_inline_field_topology_changed", Assert.Single(topologyRejected.Diagnostics).Code);

        paragraph.Paragraph.Runs[1].InlineField.Instruction = "SEQ Figure \\* ROMAN";
        var invalid = Invoke(new CodecRequest
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
        });
        Assert.False(invalid.Ok);
        Assert.Equal("invalid_document_inline_field", Assert.Single(invalid.Diagnostics).Code);
    }

    [Fact]
    public void InlinePlainTextContentControlAuthorsImportsEditsAndRejectsComplexTopology()
    {
        var document = new DocumentArtifact { Id = "document/content-controls", Name = "Content-control template" };
        var paragraph = new DocumentBlock
        {
            Id = "document/content-controls/customer",
            StyleId = "Normal",
            Paragraph = new DocumentParagraph { Text = "Customer: Ada Lovelace." },
        };
        paragraph.Paragraph.Runs.Add(new DocumentRun { Text = "Customer: " });
        paragraph.Paragraph.Runs.Add(new DocumentRun
        {
            Text = "Ada Lovelace",
            TextContentControl = new DocumentTextContentControl
            {
                Id = "customer-name",
                Tag = "CUSTOMER_NAME",
                Alias = "Customer name",
            },
        });
        paragraph.Paragraph.Runs.Add(new DocumentRun { Text = "." });
        document.Blocks.Add(paragraph);

        var authored = Invoke(new CodecRequest
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
        });
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var package = WordprocessingDocument.Open(stream, false))
        {
            var control = Assert.Single(package.MainDocumentPart!.Document!.Descendants<W.SdtRun>());
            Assert.Equal("CUSTOMER_NAME", control.SdtProperties!.GetFirstChild<W.Tag>()!.Val!.Value);
            Assert.Equal("Customer name", control.SdtProperties.GetFirstChild<W.SdtAlias>()!.Val!.Value);
            Assert.True(control.SdtProperties.GetFirstChild<W.SdtId>()!.Val!.Value > 0);
            Assert.NotNull(control.SdtProperties.GetFirstChild<W.SdtContentText>());
            Assert.Equal("Ada Lovelace", control.InnerText);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var importedParagraph = Assert.Single(imported.Artifact.Document.Blocks).Paragraph;
        var importedControlRun = Assert.Single(importedParagraph.Runs, run => run.TextContentControl is not null);
        Assert.Equal("CUSTOMER_NAME", importedControlRun.TextContentControl.Tag);
        Assert.True(importedControlRun.TextContentControl.HasNativeId);

        var unchanged = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact.Clone(),
        });
        Assert.True(unchanged.Ok, Diagnostics(unchanged));
        Assert.Equal(authored.File, unchanged.File);

        importedControlRun.Text = "Grace Hopper";
        importedControlRun.TextContentControl.Tag = "CONTACT_NAME";
        importedControlRun.TextContentControl.Alias = "Contact name";
        importedParagraph.Text = "Customer: Grace Hopper.";
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        var roundTrip = DocxCodec.Import(edited.File.ToByteArray(), EffectiveCodecLimits.From(null)).Artifact;
        var editedRun = Assert.Single(Assert.Single(roundTrip.Document.Blocks).Paragraph.Runs, run => run.TextContentControl is not null);
        Assert.Equal("Grace Hopper", editedRun.Text);
        Assert.Equal("CONTACT_NAME", editedRun.TextContentControl.Tag);
        Assert.Equal("Contact name", editedRun.TextContentControl.Alias);

        var changedTopology = roundTrip.Clone();
        Assert.Single(changedTopology.Document.Blocks).Paragraph.Runs[1].TextContentControl = null;
        var rejectedTopology = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = changedTopology,
        });
        Assert.False(rejectedTopology.Ok);
        Assert.Equal("document_content_control_topology_changed", Assert.Single(rejectedTopology.Diagnostics).Code);

        var complexBytes = AddDatePickerContentControl(authored.File.ToByteArray());
        var complex = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(complexBytes),
        });
        Assert.True(complex.Ok, Diagnostics(complex));
        var preserved = Assert.Single(complex.Artifact.Document.Blocks);
        Assert.False(preserved.Source.Editable);
        Assert.Empty(preserved.Paragraph.Runs);
        var preservedExport = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = complex.Artifact,
        });
        Assert.True(preservedExport.Ok, Diagnostics(preservedExport));
        Assert.Equal(ByteString.CopyFrom(complexBytes), preservedExport.File);
        preserved.Paragraph.Text = "Unsafe date-picker edit";
        var rejectedComplex = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = complex.Artifact,
        });
        Assert.False(rejectedComplex.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(rejectedComplex.Diagnostics).Code);
    }

    [Fact]
    public void BlockPlainTextContentControlAuthorsImportsEditsAndRejectsComplexTopology()
    {
        var document = new DocumentArtifact { Id = "document/block-content-control", Name = "Block content-control template" };
        var paragraph = new DocumentBlock
        {
            Id = "document/block-content-control/summary",
            StyleId = "Normal",
            Paragraph = new DocumentParagraph
            {
                Text = "Executive summary",
                Formatting = new DocumentParagraphFormatting { KeepNext = true },
                BlockContentControl = new DocumentTextContentControl
                {
                    Id = "executive-summary",
                    Tag = "EXECUTIVE_SUMMARY",
                    Alias = "Executive summary",
                    ControlType = DocumentContentControlType.PlainText,
                },
            },
        };
        paragraph.Paragraph.Runs.Add(new DocumentRun
        {
            Text = "Executive summary",
            Formatting = new DocumentRunFormatting { Bold = true, ColorRgb = "1D4ED8" },
        });
        document.Blocks.Add(paragraph);

        var authored = Invoke(new CodecRequest
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
        });
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var package = WordprocessingDocument.Open(stream, false))
        {
            var body = package.MainDocumentPart!.Document!.Body!;
            var control = Assert.IsType<W.SdtBlock>(body.ChildElements[0]);
            Assert.Equal("EXECUTIVE_SUMMARY", control.SdtProperties!.GetFirstChild<W.Tag>()!.Val!.Value);
            Assert.Equal("Executive summary", control.SdtProperties.GetFirstChild<W.SdtAlias>()!.Val!.Value);
            Assert.True(control.SdtProperties.GetFirstChild<W.SdtId>()!.Val!.Value > 0);
            Assert.NotNull(control.SdtProperties.GetFirstChild<W.SdtContentText>());
            var contentParagraph = Assert.Single(control.SdtContentBlock!.Elements<W.Paragraph>());
            Assert.Single(contentParagraph.Elements<W.Run>());
            Assert.Equal("Executive summary", contentParagraph.InnerText);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var importedBlock = Assert.Single(imported.Artifact.Document.Blocks);
        Assert.True(importedBlock.Source.Editable);
        var importedParagraph = importedBlock.Paragraph;
        Assert.Equal("EXECUTIVE_SUMMARY", importedParagraph.BlockContentControl.Tag);
        Assert.Equal(DocumentContentControlType.PlainText, importedParagraph.BlockContentControl.ControlType);
        Assert.True(importedParagraph.BlockContentControl.HasNativeId);
        Assert.True(Assert.Single(importedParagraph.Runs).Formatting.Bold);

        var unchanged = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact.Clone(),
        });
        Assert.True(unchanged.Ok, Diagnostics(unchanged));
        Assert.Equal(authored.File, unchanged.File);

        importedParagraph.Text = "Updated executive summary";
        importedParagraph.Runs[0].Text = importedParagraph.Text;
        importedParagraph.BlockContentControl.Tag = "SUMMARY";
        importedParagraph.BlockContentControl.Alias = "Summary";
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        var roundTrip = DocxCodec.Import(edited.File.ToByteArray(), EffectiveCodecLimits.From(null)).Artifact;
        var editedParagraph = Assert.Single(roundTrip.Document.Blocks).Paragraph;
        Assert.Equal("Updated executive summary", editedParagraph.Text);
        Assert.Equal("SUMMARY", editedParagraph.BlockContentControl.Tag);
        Assert.Equal("Summary", editedParagraph.BlockContentControl.Alias);

        var emptyAlias = roundTrip.Clone();
        Assert.Single(emptyAlias.Document.Blocks).Paragraph.BlockContentControl.Alias = string.Empty;
        var rejectedAlias = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = emptyAlias,
        });
        Assert.False(rejectedAlias.Ok);
        Assert.Equal("invalid_document_content_control", Assert.Single(rejectedAlias.Diagnostics).Code);

        var changedTopology = roundTrip.Clone();
        Assert.Single(changedTopology.Document.Blocks).Paragraph.BlockContentControl = null;
        var rejectedTopology = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = changedTopology,
        });
        Assert.False(rejectedTopology.Ok);
        Assert.Equal("document_content_control_topology_changed", Assert.Single(rejectedTopology.Diagnostics).Code);

        var complexBytes = AddSecondParagraphToBlockContentControl(authored.File.ToByteArray());
        var complex = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(complexBytes),
        });
        Assert.True(complex.Ok, Diagnostics(complex));
        var preserved = Assert.Single(complex.Artifact.Document.Blocks);
        Assert.Equal(DocumentBlock.ContentOneofCase.Opaque, preserved.ContentCase);
        Assert.False(preserved.Source.Editable);
        var preservedExport = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = complex.Artifact,
        });
        Assert.True(preservedExport.Ok, Diagnostics(preservedExport));
        Assert.Equal(ByteString.CopyFrom(complexBytes), preservedExport.File);
        preserved.Opaque.Text = "Unsafe multi-paragraph edit";
        var rejectedComplex = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = complex.Artifact,
        });
        Assert.False(rejectedComplex.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(rejectedComplex.Diagnostics).Code);
    }

    [Fact]
    public void InlineCheckboxContentControlAuthorsImportsEditsAndRejectsIrregularGraphs()
    {
        var document = new DocumentArtifact { Id = "document/checkbox", Name = "Checkbox template" };
        var paragraph = new DocumentBlock
        {
            Id = "document/checkbox/terms",
            StyleId = "Normal",
            Paragraph = new DocumentParagraph { Text = "Accept terms: ☐ I agree." },
        };
        paragraph.Paragraph.Runs.Add(new DocumentRun { Text = "Accept terms: " });
        paragraph.Paragraph.Runs.Add(new DocumentRun
        {
            Text = "☐",
            TextContentControl = new DocumentTextContentControl
            {
                Id = "terms-accepted",
                Tag = "TERMS_ACCEPTED",
                Alias = "Terms accepted",
                ControlType = DocumentContentControlType.Checkbox,
                Checked = false,
            },
        });
        paragraph.Paragraph.Runs.Add(new DocumentRun { Text = " I agree." });
        document.Blocks.Add(paragraph);

        var authored = Invoke(new CodecRequest
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
        });
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var package = WordprocessingDocument.Open(stream, false))
        {
            var control = Assert.Single(package.MainDocumentPart!.Document!.Descendants<W.SdtRun>());
            var checkbox = Assert.IsType<W14.SdtContentCheckBox>(control.SdtProperties!.LastChild);
            Assert.Equal(W14.OnOffValues.Zero, checkbox.GetFirstChild<W14.Checked>()!.Val!.Value);
            Assert.Equal("2612", checkbox.GetFirstChild<W14.CheckedState>()!.Val!.Value);
            Assert.Equal("MS Gothic", checkbox.GetFirstChild<W14.CheckedState>()!.Font!.Value);
            Assert.Equal("2610", checkbox.GetFirstChild<W14.UncheckedState>()!.Val!.Value);
            Assert.Equal("☐", control.InnerText);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var importedParagraph = Assert.Single(imported.Artifact.Document.Blocks).Paragraph;
        var importedRun = Assert.Single(importedParagraph.Runs, run => run.TextContentControl is not null);
        Assert.Equal(DocumentContentControlType.Checkbox, importedRun.TextContentControl.ControlType);
        Assert.False(importedRun.TextContentControl.Checked);

        var unchanged = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact.Clone(),
        });
        Assert.True(unchanged.Ok, Diagnostics(unchanged));
        Assert.Equal(authored.File, unchanged.File);

        importedRun.TextContentControl.Checked = true;
        importedRun.Text = "☒";
        importedParagraph.Text = "Accept terms: ☒ I agree.";
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        var roundTrip = DocxCodec.Import(edited.File.ToByteArray(), EffectiveCodecLimits.From(null)).Artifact;
        var editedRun = Assert.Single(Assert.Single(roundTrip.Document.Blocks).Paragraph.Runs, run => run.TextContentControl is not null);
        Assert.True(editedRun.TextContentControl.Checked);
        Assert.Equal("☒", editedRun.Text);

        var changedType = roundTrip.Clone();
        var changedRun = Assert.Single(Assert.Single(changedType.Document.Blocks).Paragraph.Runs, run => run.TextContentControl is not null);
        changedRun.TextContentControl.ControlType = DocumentContentControlType.PlainText;
        var rejectedType = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = changedType,
        });
        Assert.False(rejectedType.Ok);
        Assert.Equal("document_content_control_topology_changed", Assert.Single(rejectedType.Diagnostics).Code);

        var irregularBytes = AddCustomCheckboxSymbol(authored.File.ToByteArray());
        var irregular = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(irregularBytes),
        });
        Assert.True(irregular.Ok, Diagnostics(irregular));
        var preserved = Assert.Single(irregular.Artifact.Document.Blocks);
        Assert.False(preserved.Source.Editable);
        Assert.Empty(preserved.Paragraph.Runs);
        var preservedExport = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = irregular.Artifact,
        });
        Assert.True(preservedExport.Ok, Diagnostics(preservedExport));
        Assert.Equal(ByteString.CopyFrom(irregularBytes), preservedExport.File);
    }

    [Fact]
    public void InlineDropdownContentControlAuthorsImportsEditsAndBindsChoiceTopology()
    {
        var document = new DocumentArtifact { Id = "document/dropdown", Name = "Drop-down template" };
        var paragraph = new DocumentBlock
        {
            Id = "document/dropdown/priority",
            StyleId = "Normal",
            Paragraph = new DocumentParagraph { Text = "Priority: Medium." },
        };
        paragraph.Paragraph.Runs.Add(new DocumentRun { Text = "Priority: " });
        var control = new DocumentTextContentControl
        {
            Id = "priority",
            Tag = "PRIORITY",
            Alias = "Priority",
            ControlType = DocumentContentControlType.DropDown,
            SelectedValue = "medium",
        };
        control.Choices.Add(new[]
        {
            new DocumentContentControlChoice { DisplayText = "Low", Value = "low" },
            new DocumentContentControlChoice { DisplayText = "Medium", Value = "medium" },
            new DocumentContentControlChoice { DisplayText = "High", Value = "high" },
        });
        paragraph.Paragraph.Runs.Add(new DocumentRun { Text = "Medium", TextContentControl = control });
        paragraph.Paragraph.Runs.Add(new DocumentRun { Text = "." });
        document.Blocks.Add(paragraph);

        var authored = Invoke(new CodecRequest
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
        });
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var package = WordprocessingDocument.Open(stream, false))
        {
            var sdt = Assert.Single(package.MainDocumentPart!.Document!.Descendants<W.SdtRun>());
            var dropdown = Assert.IsType<W.SdtContentDropDownList>(sdt.SdtProperties!.LastChild);
            Assert.Equal("medium", dropdown.LastValue!.Value);
            Assert.Equal(
                new[] { ("Low", "low"), ("Medium", "medium"), ("High", "high") },
                dropdown.Elements<W.ListItem>().Select(item => (item.DisplayText!.Value!, item.Value!.Value!)).ToArray());
            Assert.Equal("Medium", sdt.InnerText);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var importedParagraph = Assert.Single(imported.Artifact.Document.Blocks).Paragraph;
        var importedRun = Assert.Single(importedParagraph.Runs, run => run.TextContentControl is not null);
        Assert.Equal(DocumentContentControlType.DropDown, importedRun.TextContentControl.ControlType);
        Assert.Equal("medium", importedRun.TextContentControl.SelectedValue);
        Assert.Equal(new[] { "low", "medium", "high" }, importedRun.TextContentControl.Choices.Select(choice => choice.Value).ToArray());

        var unchanged = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact.Clone(),
        });
        Assert.True(unchanged.Ok, Diagnostics(unchanged));
        Assert.Equal(authored.File, unchanged.File);

        importedRun.TextContentControl.SelectedValue = "high";
        importedRun.Text = "High";
        importedParagraph.Text = "Priority: High.";
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        var roundTrip = DocxCodec.Import(edited.File.ToByteArray(), EffectiveCodecLimits.From(null)).Artifact;
        var editedRun = Assert.Single(Assert.Single(roundTrip.Document.Blocks).Paragraph.Runs, run => run.TextContentControl is not null);
        Assert.Equal("high", editedRun.TextContentControl.SelectedValue);
        Assert.Equal("High", editedRun.Text);

        var changedChoices = roundTrip.Clone();
        var changedChoiceRun = Assert.Single(Assert.Single(changedChoices.Document.Blocks).Paragraph.Runs, run => run.TextContentControl is not null);
        changedChoiceRun.TextContentControl.Choices[0].DisplayText = "Routine";
        var rejectedChoices = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = changedChoices,
        });
        Assert.False(rejectedChoices.Ok);
        Assert.Equal("document_content_control_topology_changed", Assert.Single(rejectedChoices.Diagnostics).Code);

        var invalidSelection = roundTrip.Clone();
        var invalidSelectionRun = Assert.Single(Assert.Single(invalidSelection.Document.Blocks).Paragraph.Runs, run => run.TextContentControl is not null);
        invalidSelectionRun.TextContentControl.SelectedValue = "urgent";
        var rejectedSelection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = invalidSelection,
        });
        Assert.False(rejectedSelection.Ok);
        Assert.Equal("invalid_document_content_control", Assert.Single(rejectedSelection.Diagnostics).Code);

        var irregularBytes = DuplicateDropdownDisplayText(authored.File.ToByteArray());
        var irregular = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(irregularBytes),
        });
        Assert.True(irregular.Ok, Diagnostics(irregular));
        var preserved = Assert.Single(irregular.Artifact.Document.Blocks);
        Assert.False(preserved.Source.Editable);
        Assert.Empty(preserved.Paragraph.Runs);
        var preservedExport = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = irregular.Artifact,
        });
        Assert.True(preservedExport.Ok, Diagnostics(preservedExport));
        Assert.Equal(ByteString.CopyFrom(irregularBytes), preservedExport.File);
    }

    [Fact]
    public void InlineComboBoxContentControlAuthorsImportsCustomValuesAndBindsChoiceTopology()
    {
        var document = new DocumentArtifact { Id = "document/combo-box", Name = "Combo-box template" };
        var paragraph = new DocumentBlock
        {
            Id = "document/combo-box/contact",
            StyleId = "Normal",
            Paragraph = new DocumentParagraph { Text = "Contact method: Email." },
        };
        paragraph.Paragraph.Runs.Add(new DocumentRun { Text = "Contact method: " });
        var control = new DocumentTextContentControl
        {
            Id = "contact-method",
            Tag = "CONTACT_METHOD",
            Alias = "Contact method",
            ControlType = DocumentContentControlType.ComboBox,
            Value = "email",
        };
        control.Choices.Add(new[]
        {
            new DocumentContentControlChoice { DisplayText = "Email", Value = "email" },
            new DocumentContentControlChoice { DisplayText = "Phone call", Value = "phone" },
        });
        paragraph.Paragraph.Runs.Add(new DocumentRun { Text = "Email", TextContentControl = control });
        paragraph.Paragraph.Runs.Add(new DocumentRun { Text = "." });
        document.Blocks.Add(paragraph);

        var authored = Invoke(new CodecRequest
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
        });
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var package = WordprocessingDocument.Open(stream, false))
        {
            var sdt = Assert.Single(package.MainDocumentPart!.Document!.Descendants<W.SdtRun>());
            var comboBox = Assert.IsType<W.SdtContentComboBox>(sdt.SdtProperties!.LastChild);
            Assert.Equal("email", comboBox.LastValue!.Value);
            Assert.Equal(
                new[] { ("Email", "email"), ("Phone call", "phone") },
                comboBox.Elements<W.ListItem>().Select(item => (item.DisplayText!.Value!, item.Value!.Value!)).ToArray());
            Assert.Equal("Email", sdt.InnerText);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var importedParagraph = Assert.Single(imported.Artifact.Document.Blocks).Paragraph;
        var importedRun = Assert.Single(importedParagraph.Runs, run => run.TextContentControl is not null);
        Assert.Equal(DocumentContentControlType.ComboBox, importedRun.TextContentControl.ControlType);
        Assert.Equal("email", importedRun.TextContentControl.Value);
        Assert.Equal(new[] { "email", "phone" }, importedRun.TextContentControl.Choices.Select(choice => choice.Value).ToArray());

        var unchanged = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact.Clone(),
        });
        Assert.True(unchanged.Ok, Diagnostics(unchanged));
        Assert.Equal(authored.File, unchanged.File);

        importedRun.TextContentControl.Value = "Pager duty";
        importedRun.Text = "Pager duty";
        importedParagraph.Text = "Contact method: Pager duty.";
        var custom = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(custom.Ok, Diagnostics(custom));
        var customRoundTrip = DocxCodec.Import(custom.File.ToByteArray(), EffectiveCodecLimits.From(null)).Artifact;
        var customRun = Assert.Single(Assert.Single(customRoundTrip.Document.Blocks).Paragraph.Runs, run => run.TextContentControl is not null);
        Assert.Equal("Pager duty", customRun.TextContentControl.Value);
        Assert.Equal("Pager duty", customRun.Text);

        customRun.TextContentControl.Value = "phone";
        customRun.Text = "Phone call";
        Assert.Single(customRoundTrip.Document.Blocks).Paragraph.Text = "Contact method: Phone call.";
        var declared = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = customRoundTrip,
        });
        Assert.True(declared.Ok, Diagnostics(declared));
        var declaredRoundTrip = DocxCodec.Import(declared.File.ToByteArray(), EffectiveCodecLimits.From(null)).Artifact;
        var declaredRun = Assert.Single(Assert.Single(declaredRoundTrip.Document.Blocks).Paragraph.Runs, run => run.TextContentControl is not null);
        Assert.Equal("phone", declaredRun.TextContentControl.Value);
        Assert.Equal("Phone call", declaredRun.Text);

        var changedChoices = declaredRoundTrip.Clone();
        var changedChoiceRun = Assert.Single(Assert.Single(changedChoices.Document.Blocks).Paragraph.Runs, run => run.TextContentControl is not null);
        changedChoiceRun.TextContentControl.Choices[0].DisplayText = "Electronic mail";
        var rejectedChoices = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = changedChoices,
        });
        Assert.False(rejectedChoices.Ok);
        Assert.Equal("document_content_control_topology_changed", Assert.Single(rejectedChoices.Diagnostics).Code);

        var invalidVisibleText = declaredRoundTrip.Clone();
        var invalidRun = Assert.Single(Assert.Single(invalidVisibleText.Document.Blocks).Paragraph.Runs, run => run.TextContentControl is not null);
        invalidRun.Text = "phone";
        var rejectedVisibleText = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = invalidVisibleText,
        });
        Assert.False(rejectedVisibleText.Ok);
        Assert.Equal("invalid_document_content_control", Assert.Single(rejectedVisibleText.Diagnostics).Code);
    }

    [Fact]
    public void InlineDateContentControlAuthorsImportsEditsAndRejectsNoncanonicalProfiles()
    {
        var document = new DocumentArtifact { Id = "document/date", Name = "Date template" };
        var paragraph = new DocumentBlock
        {
            Id = "document/date/review",
            StyleId = "Normal",
            Paragraph = new DocumentParagraph { Text = "Review date: 2026-07-21." },
        };
        paragraph.Paragraph.Runs.Add(new DocumentRun { Text = "Review date: " });
        paragraph.Paragraph.Runs.Add(new DocumentRun
        {
            Text = "2026-07-21",
            TextContentControl = new DocumentTextContentControl
            {
                Id = "review-date",
                Tag = "REVIEW_DATE",
                Alias = "Review date",
                ControlType = DocumentContentControlType.Date,
                DateValue = "2026-07-21",
            },
        });
        paragraph.Paragraph.Runs.Add(new DocumentRun { Text = "." });
        document.Blocks.Add(paragraph);

        var authored = Invoke(new CodecRequest
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
        });
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var package = WordprocessingDocument.Open(stream, false))
        {
            var sdt = Assert.Single(package.MainDocumentPart!.Document!.Descendants<W.SdtRun>());
            var date = Assert.IsType<W.SdtContentDate>(sdt.SdtProperties!.LastChild);
            Assert.Equal("2026-07-21T00:00:00Z", date.FullDate!.InnerText);
            Assert.Equal("yyyy-MM-dd", date.DateFormat!.Val!.Value);
            Assert.Equal("en-US", date.LanguageId!.Val!.Value);
            Assert.Equal(W.DateFormatValues.Date, date.SdtDateMappingType!.Val!.Value);
            Assert.Equal(W.CalendarValues.Gregorian, date.Calendar!.Val!.Value);
            Assert.Equal("2026-07-21", sdt.InnerText);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var importedParagraph = Assert.Single(imported.Artifact.Document.Blocks).Paragraph;
        var importedRun = Assert.Single(importedParagraph.Runs, run => run.TextContentControl is not null);
        Assert.Equal(DocumentContentControlType.Date, importedRun.TextContentControl.ControlType);
        Assert.Equal("2026-07-21", importedRun.TextContentControl.DateValue);

        var unchanged = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact.Clone(),
        });
        Assert.True(unchanged.Ok, Diagnostics(unchanged));
        Assert.Equal(authored.File, unchanged.File);

        importedRun.TextContentControl.DateValue = "2028-02-29";
        importedRun.Text = "2028-02-29";
        importedParagraph.Text = "Review date: 2028-02-29.";
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        var roundTrip = DocxCodec.Import(edited.File.ToByteArray(), EffectiveCodecLimits.From(null)).Artifact;
        var editedRun = Assert.Single(Assert.Single(roundTrip.Document.Blocks).Paragraph.Runs, run => run.TextContentControl is not null);
        Assert.Equal("2028-02-29", editedRun.TextContentControl.DateValue);
        Assert.Equal("2028-02-29", editedRun.Text);

        var changedType = roundTrip.Clone();
        Assert.Single(Assert.Single(changedType.Document.Blocks).Paragraph.Runs, run => run.TextContentControl is not null).TextContentControl.ControlType = DocumentContentControlType.PlainText;
        var rejectedType = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = changedType,
        });
        Assert.False(rejectedType.Ok);
        Assert.Equal("document_content_control_topology_changed", Assert.Single(rejectedType.Diagnostics).Code);

        var invalidDate = roundTrip.Clone();
        var invalidDateRun = Assert.Single(Assert.Single(invalidDate.Document.Blocks).Paragraph.Runs, run => run.TextContentControl is not null);
        invalidDateRun.TextContentControl.DateValue = "2026-02-29";
        invalidDateRun.Text = "2026-02-29";
        Assert.Single(invalidDate.Document.Blocks).Paragraph.Text = "Review date: 2026-02-29.";
        var rejectedDate = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = invalidDate,
        });
        Assert.False(rejectedDate.Ok);
        Assert.Equal("invalid_document_content_control", Assert.Single(rejectedDate.Diagnostics).Code);

        var noncanonicalBytes = ChangeDatePickerFormat(authored.File.ToByteArray(), "M/d/yyyy");
        var noncanonical = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(noncanonicalBytes),
        });
        Assert.True(noncanonical.Ok, Diagnostics(noncanonical));
        var preserved = Assert.Single(noncanonical.Artifact.Document.Blocks);
        Assert.False(preserved.Source.Editable);
        Assert.Empty(preserved.Paragraph.Runs);
        var preservedExport = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = noncanonical.Artifact,
        });
        Assert.True(preservedExport.Ok, Diagnostics(preservedExport));
        Assert.Equal(ByteString.CopyFrom(noncanonicalBytes), preservedExport.File);
    }

    [Fact]
    public void OfficeSkillProfileRoundTripsFormattingImagesSectionsAndHeaders()
    {
        var authored = Invoke(OfficeSkillProfileExportRequest());
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var package = WordprocessingDocument.Open(stream, false))
        {
            var mainPart = package.MainDocumentPart!;
            Assert.NotNull(mainPart.StyleDefinitionsPart?.Styles?.DocDefaults?.RunPropertiesDefault);
            Assert.Contains(mainPart.StyleDefinitionsPart!.Styles!.Elements<W.Style>(), style => style.StyleId == "BriefLead");
            Assert.Single(mainPart.ImageParts);
            Assert.Equal(2, mainPart.HeaderParts.Count());
            Assert.Single(mainPart.FooterParts);
            Assert.NotNull(mainPart.DocumentSettingsPart?.Settings?.GetFirstChild<W.EvenAndOddHeaders>());
            Assert.Equal(2, mainPart.Document!.Body!.Descendants<W.SectionProperties>().Count());
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var document = imported.Artifact.Document;
        Assert.Equal("Aptos", document.DefaultRunStyle.FontFamily);
        Assert.Contains(document.Styles, style => style.Id == "BriefLead" && style.BasedOn == "Normal");
        Assert.True(document.EvenAndOddHeaders);
        Assert.Equal(2, document.Headers.Count);
        Assert.Single(document.Footers);
        Assert.True(Assert.Single(document.SectionSettings).DifferentFirstPage);

        var paragraph = document.Blocks.Single(block => block.ContentCase == DocumentBlock.ContentOneofCase.Paragraph && block.StyleId == "BriefLead");
        Assert.True(paragraph.Source.Editable);
        Assert.Equal("center", paragraph.Paragraph.Formatting.Alignment);
        Assert.True(paragraph.Paragraph.Formatting.KeepNext);
        Assert.Equal("Aptos Display", paragraph.Paragraph.Runs[0].Formatting.FontFamily);
        Assert.Equal((uint)30, paragraph.Paragraph.Runs[0].Formatting.FontSizeHalfPoints);
        Assert.Equal("315A83", paragraph.Paragraph.Runs[0].Formatting.ColorRgb);
        Assert.True(paragraph.Paragraph.Runs[0].Formatting.Bold);
        Assert.True(paragraph.Paragraph.Runs[0].Formatting.Underline);
        Assert.True(paragraph.Paragraph.Runs[0].Formatting.HasItalic);
        Assert.False(paragraph.Paragraph.Runs[0].Formatting.Italic);
        Assert.Contains(document.Blocks, block => block.ContentCase == DocumentBlock.ContentOneofCase.Field && block.Field.Instruction == "PAGE");
        var image = document.Blocks.Single(block => block.ContentCase == DocumentBlock.ContentOneofCase.Image);
        Assert.True(image.Source.Editable);
        Assert.Single(imported.Artifact.Assets);
        var section = document.Blocks.Single(block => block.ContentCase == DocumentBlock.ContentOneofCase.Section);
        Assert.True(section.Section.Landscape);
        Assert.Equal(DocumentSectionBreak.Continuous, section.Section.BreakType);

        paragraph.Paragraph.Text = "Edited styled lead";
        paragraph.Paragraph.Runs[0].Text = paragraph.Paragraph.Text;
        paragraph.Paragraph.Runs[0].Formatting.ColorRgb = "9C2B2E";
        image.Image.AltText = "Edited chart preview";
        image.Image.WidthEmu += 9_525;
        section.Section.MarginLeftTwips = 1_200;
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        using (var stream = new MemoryStream(edited.File.ToByteArray()))
        using (var package = WordprocessingDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));

        var roundTrip = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = edited.File,
        });
        Assert.True(roundTrip.Ok, Diagnostics(roundTrip));
        Assert.Contains(roundTrip.Artifact.Document.Blocks, block =>
            block.ContentCase == DocumentBlock.ContentOneofCase.Paragraph && block.Paragraph.Text == "Edited styled lead" &&
            block.Paragraph.Runs[0].Formatting.ColorRgb == "9C2B2E");
        Assert.Contains(roundTrip.Artifact.Document.Blocks, block =>
            block.ContentCase == DocumentBlock.ContentOneofCase.Image && block.Image.AltText == "Edited chart preview");
        Assert.Equal((uint)1_200, roundTrip.Artifact.Document.Blocks.Single(block =>
            block.ContentCase == DocumentBlock.ContentOneofCase.Section).Section.MarginLeftTwips);
    }

    [Fact]
    public void OfficeSkillProfileRejectsMissingImageAssetsAndInvalidSectionGeometry()
    {
        var missingAsset = OfficeSkillProfileExportRequest();
        missingAsset.Artifact.Assets.Clear();
        var missingAssetResponse = Invoke(missingAsset);
        Assert.False(missingAssetResponse.Ok);
        Assert.Equal("invalid_document_image_asset", Assert.Single(missingAssetResponse.Diagnostics).Code);

        var invalidSection = OfficeSkillProfileExportRequest();
        var section = invalidSection.Artifact.Document.Blocks.Single(block =>
            block.ContentCase == DocumentBlock.ContentOneofCase.Section).Section;
        section.MarginLeftTwips = section.PageWidthTwips;
        var invalidSectionResponse = Invoke(invalidSection);
        Assert.False(invalidSectionResponse.Ok);
        Assert.Equal("invalid_document_section", Assert.Single(invalidSectionResponse.Diagnostics).Code);
    }

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
    public void ClassicCommentsAuthorImportAndEditFixedTopology()
    {
        var authored = Invoke(ClassicCommentExportRequest());
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var package = WordprocessingDocument.Open(stream, false))
        {
            var mainPart = package.MainDocumentPart!;
            var comments = mainPart.WordprocessingCommentsPart!.Comments!.Elements<W.Comment>().ToArray();
            Assert.Equal(2, comments.Length);
            Assert.Equal("Review this paragraph.", comments[0].InnerText);
            Assert.Equal("2026-07-16T08:00:00Z", comments[0].GetAttribute("date", "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value);
            var paragraph = mainPart.Document!.Body!.Elements<W.Paragraph>().Single();
            Assert.Equal(2, paragraph.Elements<W.CommentRangeStart>().Count());
            Assert.Equal(2, paragraph.Elements<W.CommentRangeEnd>().Count());
            Assert.Equal(2, paragraph.Descendants<W.CommentReference>().Count());
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        Assert.False(imported.Artifact.Document.Blocks[0].Source.Editable);
        Assert.Equal(
            imported.Artifact.Document.Comments[0].Source.ResidualSha256,
            imported.Artifact.Document.Comments[1].Source.ResidualSha256);
        Assert.Collection(imported.Artifact.Document.Comments,
            comment =>
            {
                Assert.Equal("document/comment/1", comment.Id);
                Assert.Equal("document/block/1", comment.TargetBlockId);
                Assert.Equal("Reviewer", comment.Author);
                Assert.Equal("RV", comment.Initials);
                Assert.Equal("2026-07-16T08:00:00Z", comment.CreatedAt);
                Assert.True(comment.Source.Editable);
                Assert.Equal("0", comment.Source.NativeCommentId);
                Assert.NotEmpty(comment.Source.CommentElementSha256);
                Assert.NotEmpty(comment.Source.SemanticSha256);
                Assert.NotEmpty(comment.Source.ResidualSha256);
                Assert.NotEmpty(comment.Source.AnchorSha256);
            },
            comment =>
            {
                Assert.Equal("Second reviewer", comment.Author);
                Assert.False(comment.HasCreatedAt);
            });

        var edited = imported.Artifact.Document.Comments[0];
        edited.Author = "Lead reviewer";
        edited.Text = "Approved after source-bound review.";
        edited.ClearInitials();
        edited.CreatedAt = "2026-07-16T09:30:00+08:00";
        var exported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(exported.Ok, Diagnostics(exported));
        Assert.Equal("opaque_content_preserved", Assert.Single(exported.Diagnostics).Code);
        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var package = WordprocessingDocument.Open(stream, false))
        {
            var comments = package.MainDocumentPart!.WordprocessingCommentsPart!.Comments!.Elements<W.Comment>().ToArray();
            Assert.Equal("Lead reviewer", comments[0].Author?.Value);
            Assert.Null(comments[0].Initials);
            Assert.Equal("2026-07-16T09:30:00+08:00", comments[0].GetAttribute("date", "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value);
            Assert.Equal("Approved after source-bound review.", comments[0].InnerText);
            Assert.Equal("Keep the evidence link.", comments[1].InnerText);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }
        var reimported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = exported.File,
        });
        Assert.True(reimported.Ok, Diagnostics(reimported));
        Assert.Equal("Approved after source-bound review.", reimported.Artifact.Document.Comments[0].Text);
        Assert.False(reimported.Artifact.Document.Comments[0].HasInitials);
    }

    [Fact]
    public void ClassicCommentsFailClosedOnTopologyBindingAndComplexBodies()
    {
        var authored = Invoke(ClassicCommentExportRequest());
        Assert.True(authored.Ok, Diagnostics(authored));
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));

        var removed = imported.Artifact.Clone();
        removed.Document.Comments.RemoveAt(1);
        var topology = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = removed,
        });
        Assert.False(topology.Ok);
        Assert.Equal("document_comment_topology_changed", Assert.Single(topology.Diagnostics).Code);

        var tampered = imported.Artifact.Clone();
        tampered.Document.Comments[0].Source.AnchorSha256 = new string('0', 64);
        var binding = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = tampered,
        });
        Assert.False(binding.Ok);
        Assert.Equal("document_comment_source_binding_mismatch", Assert.Single(binding.Diagnostics).Code);

        var complex = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(AddSecondCommentBodyRun(authored.File.ToByteArray())),
        });
        Assert.True(complex.Ok, Diagnostics(complex));
        Assert.Empty(complex.Artifact.Document.Comments);
        Assert.Contains(complex.Diagnostics, item => item.Code == "unsupported_document_comments_preserved");
        var unchanged = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = complex.Artifact,
        });
        Assert.True(unchanged.Ok, Diagnostics(unchanged));
        Assert.Equal(complex.Artifact.OpaqueOpc.SourcePackage.Data.ToByteArray(), unchanged.File.ToByteArray());

        var aliasedIds = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(AliasSecondCommentId(authored.File.ToByteArray())),
        });
        Assert.True(aliasedIds.Ok, Diagnostics(aliasedIds));
        Assert.Empty(aliasedIds.Artifact.Document.Comments);
        Assert.Contains(aliasedIds.Diagnostics, item => item.Code == "unsupported_document_comments_preserved");
    }

    [Fact]
    public void ModernCommentThreadAuthorsImportsAndEditsTextAndResolvedState()
    {
        var authored = Invoke(ModernCommentExportRequest());
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var package = WordprocessingDocument.Open(stream, false))
        {
            var main = package.MainDocumentPart!;
            Assert.Equal(2, main.WordprocessingCommentsPart!.Comments!.Elements<W.Comment>().Count());
            Assert.Single(main.Document!.Body!.Descendants<W.CommentRangeStart>());
            Assert.Single(main.Document.Body.Descendants<W.CommentRangeEnd>());
            Assert.Single(main.Document.Body.Descendants<W.CommentReference>());
            var extended = main.WordprocessingCommentsExPart!.CommentsEx!.Elements<W15.CommentEx>().ToArray();
            Assert.Equal(2, extended.Length);
            Assert.Equal("11111111", extended[0].GetAttribute("paraId", "http://schemas.microsoft.com/office/word/2012/wordml").Value);
            Assert.Equal("11111111", extended[1].GetAttribute("paraIdParent", "http://schemas.microsoft.com/office/word/2012/wordml").Value);
            Assert.Equal(2, main.WordprocessingCommentsIdsPart!.CommentsIds!.Elements<W16Cid.CommentId>().Count());
            Assert.Equal(2, main.WordCommentsExtensiblePart!.CommentsExtensible!.Elements<W16Cex.CommentExtensible>().Count());
            Assert.Equal(2, main.WordprocessingPeoplePart!.People!.Elements<W15.Person>().Count());
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        Assert.Collection(imported.Artifact.Document.Comments,
            root =>
            {
                Assert.Empty(root.ParentCommentId);
                Assert.False(root.Resolved);
                Assert.Equal("11111111", root.ParagraphId);
                Assert.Equal("33333333", root.DurableId);
                Assert.Equal("2026-07-19T08:00:00Z", root.DateUtc);
                Assert.Equal("provider-a", root.Person.ProviderId);
                Assert.True(root.Source.ThreadEditable);
                Assert.NotEmpty(root.Source.ExtendedGraphSha256);
                Assert.NotEmpty(root.Source.CommentsIdsGraphSha256);
                Assert.NotEmpty(root.Source.CommentsExtensibleGraphSha256);
                Assert.NotEmpty(root.Source.PeopleGraphSha256);
            },
            reply =>
            {
                Assert.Equal("document/comment/1", reply.ParentCommentId);
                Assert.False(reply.Resolved);
                Assert.Equal("22222222", reply.ParagraphId);
                Assert.Equal("44444444", reply.DurableId);
                Assert.Equal("provider-b", reply.Person.ProviderId);
                Assert.Equal(imported.Artifact.Document.Comments[0].Source.AnchorSha256, reply.Source.AnchorSha256);
            });

        imported.Artifact.Document.Comments[0].Text = "Resolved after bounded review.";
        imported.Artifact.Document.Comments[0].Resolved = true;
        imported.Artifact.Document.Comments[1].Text = "Reply retained with its root.";
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        using (var stream = new MemoryStream(edited.File.ToByteArray()))
        using (var package = WordprocessingDocument.Open(stream, false))
        {
            var extended = package.MainDocumentPart!.WordprocessingCommentsExPart!.CommentsEx!.Elements<W15.CommentEx>().ToArray();
            Assert.Equal("1", extended[0].GetAttribute("done", "http://schemas.microsoft.com/office/word/2012/wordml").Value);
            Assert.Equal("11111111", extended[1].GetAttribute("paraIdParent", "http://schemas.microsoft.com/office/word/2012/wordml").Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }
        var reimported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = edited.File,
        });
        Assert.True(reimported.Ok, Diagnostics(reimported));
        Assert.Equal("Resolved after bounded review.", reimported.Artifact.Document.Comments[0].Text);
        Assert.True(reimported.Artifact.Document.Comments[0].Resolved);
        Assert.Equal("Reply retained with its root.", reimported.Artifact.Document.Comments[1].Text);
        Assert.Equal("document/comment/1", reimported.Artifact.Document.Comments[1].ParentCommentId);
    }

    [Fact]
    public void ModernCommentGeneratedParagraphIdsStayInsideTheOpenXmlRange()
    {
        var request = ModernCommentExportRequest();
        foreach (var comment in request.Artifact.Document.Comments)
        {
            comment.ParagraphId = string.Empty;
            comment.DurableId = string.Empty;
        }

        var authored = Invoke(request);
        Assert.True(authored.Ok, Diagnostics(authored));
        using var stream = new MemoryStream(authored.File.ToByteArray());
        using var package = WordprocessingDocument.Open(stream, false);
        var paragraphIds = package.MainDocumentPart!.WordprocessingCommentsExPart!.CommentsEx!
            .Elements<W15.CommentEx>()
            .Select(element => element.GetAttribute("paraId", "http://schemas.microsoft.com/office/word/2012/wordml").Value)
            .ToArray();
        Assert.Equal(2, paragraphIds.Length);
        Assert.All(paragraphIds, value =>
        {
            var number = Convert.ToUInt32(value, 16);
            Assert.True(number > 0 && number < 0x80000000, $"Generated paragraph ID {value} is outside the Open XML range.");
        });
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
    }

    [Fact]
    public void ModernCommentThreadRejectsNestedRepliesAndSourceMetadataTampering()
    {
        var nested = ModernCommentExportRequest();
        nested.Artifact.Document.Comments.Add(new DocumentComment
        {
            Id = "comment/nested",
            TargetBlockId = "document/paragraph",
            ParentCommentId = "comment/reply",
            Author = "Nested reviewer",
            Text = "Nested replies are outside the bounded profile.",
        });
        var nestedResult = Invoke(nested);
        Assert.False(nestedResult.Ok);
        Assert.Equal("unsupported_document_comment_thread", Assert.Single(nestedResult.Diagnostics).Code);

        var authored = Invoke(ModernCommentExportRequest());
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));

        var tamperedBinding = imported.Artifact.Clone();
        tamperedBinding.Document.Comments[0].Source.ExtendedGraphSha256 = new string('0', 64);
        var bindingResult = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = tamperedBinding,
        });
        Assert.False(bindingResult.Ok);
        Assert.Equal("document_comment_source_binding_mismatch", Assert.Single(bindingResult.Diagnostics).Code);

        var personEdit = imported.Artifact.Clone();
        personEdit.Document.Comments[0].Person.UserId = "changed-user";
        var personResult = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = personEdit,
        });
        Assert.False(personResult.Ok);
        Assert.Equal("unsupported_document_comment_edit", Assert.Single(personResult.Diagnostics).Code);

        var irregularBytes = MakeModernReplyNested(authored.File.ToByteArray());
        var irregular = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(irregularBytes),
        });
        Assert.True(irregular.Ok, Diagnostics(irregular));
        Assert.Empty(irregular.Artifact.Document.Comments);
        Assert.Contains(irregular.Diagnostics, item => item.Code == "unsupported_document_comments_preserved");
        var preserved = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = irregular.Artifact,
        });
        Assert.True(preserved.Ok, Diagnostics(preserved));
        Assert.Equal(irregularBytes, preserved.File.ToByteArray());
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
        Assert.True(complex.Artifact.Document.Blocks[1].Source.Editable);
        Assert.False(complex.Artifact.Document.Blocks[1].Table.Rows[0].RichCells[0].Editable);
        Assert.True(complex.Artifact.Document.Blocks[1].Table.Rows[0].RichCells[0].TextPatchable);
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
    public void SourcePreservingExportPatchesOnePlainTextNodeInComplexTableCell()
    {
        var authored = Invoke(ExportRequest());
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(AddSecondTableParagraph(authored.File.ToByteArray())),
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var table = imported.Artifact.Document.Blocks[1].Table;
        Assert.Equal("Revenue detail", table.Rows[0].Cells[0]);
        Assert.False(table.Rows[0].RichCells[0].Editable);
        Assert.True(table.Rows[0].RichCells[0].TextPatchable);
        table.TextPatches.Add(new DocumentTableTextPatch
        {
            Row = 0,
            Column = 0,
            Search = " detail",
            Replacement = " detail updated",
            SourceTextSha256 = Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes("Revenue detail"))).ToLowerInvariant(),
        });

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
            var cell = document.MainDocumentPart!.Document!.Body!.Elements<W.Table>().Single().Descendants<W.TableCell>().First();
            Assert.Equal(2, cell.Elements<W.Paragraph>().Count());
            Assert.Equal("Revenue", cell.Elements<W.Paragraph>().First().InnerText);
            Assert.Equal(" detail updated", cell.Elements<W.Paragraph>().Last().InnerText);
        }
        var roundTrip = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = exported.File,
        });
        Assert.True(roundTrip.Ok, Diagnostics(roundTrip));
        Assert.Equal("Revenue detail updated", roundTrip.Artifact.Document.Blocks[1].Table.Rows[0].Cells[0]);
        Assert.True(roundTrip.Artifact.Document.Blocks[1].Table.Rows[0].RichCells[0].TextPatchable);

        var ambiguous = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(AddSecondTableParagraph(authored.File.ToByteArray())),
        });
        ambiguous.Artifact.Document.Blocks[1].Table.TextPatches.Add(new DocumentTableTextPatch
        {
            Row = 0,
            Column = 0,
            Search = "e",
            Replacement = "E",
            SourceTextSha256 = Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes("Revenue detail"))).ToLowerInvariant(),
        });
        var ambiguousRejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = ambiguous.Artifact,
        });
        Assert.False(ambiguousRejected.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(ambiguousRejected.Diagnostics).Code);

        var invalidTable = imported.Artifact.Document.Blocks[1].Table.Clone();
        invalidTable.TextPatches.Clear();
        invalidTable.TextPatches.Add(new DocumentTableTextPatch
        {
            Row = 0,
            Column = 0,
            Search = "Revenue",
            Replacement = "\ud800",
            SourceTextSha256 = Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes("Revenue detail"))).ToLowerInvariant(),
        });
        var invalidTableException = Assert.Throws<CodecException>(() => DocxTableCodec.Validate(invalidTable));
        Assert.Equal("invalid_document_table", invalidTableException.Code);
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
        Assert.Equal(DocumentTableVerticalMerge.Unspecified, horizontal.VerticalMerge);
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
    public void DirectAuthoringBuildsValidatedHorizontalAndVerticalTableMerges()
    {
        var exported = Invoke(MergedTableExportRequest());
        Assert.True(exported.Ok, Diagnostics(exported));
        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var table = document.MainDocumentPart!.Document!.Body!.Elements<W.Table>().Single();
            Assert.Equal(new[] { "3000", "3000", "3000" }, table.TableGrid!.Elements<W.GridColumn>().Select(column => column.Width!.Value));
            Assert.Equal("9000", table.TableProperties!.TableWidth!.Width!.Value);
            Assert.Equal(240, table.TableProperties.TableIndentation!.Width!.Value);
            Assert.Equal(W.TableLayoutValues.Fixed, table.TableProperties.TableLayout!.Type!.Value);
            Assert.Equal(4, table.TableProperties.TableCellMarginDefault!.ChildElements.Count);
            Assert.All(table.TableProperties.TableBorders!.ChildElements, border =>
            {
                Assert.Equal("445566", border.GetAttribute("color", "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value);
                Assert.Equal("8", border.GetAttribute("sz", "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value);
            });
            var rows = table.Elements<W.TableRow>().ToArray();
            Assert.Equal(3, rows.Length);
            var first = rows[0].Elements<W.TableCell>().ToArray();
            Assert.Equal("6000", first[0].TableCellProperties!.TableCellWidth!.Width!.Value);
            Assert.Equal(2, first[0].TableCellProperties!.GridSpan!.Val!.Value);
            Assert.Equal(W.MergedCellValues.Restart, first[0].TableCellProperties!.VerticalMerge!.Val!.Value);
            Assert.Equal("E2E8F0", first[0].TableCellProperties!.Shading!.Fill!.Value);
            Assert.NotNull(first[0].Descendants<W.Bold>().SingleOrDefault());
            Assert.Equal("Merged owner", first[0].InnerText);
            Assert.Equal(W.MergedCellValues.Continue, rows[1].Elements<W.TableCell>().First().TableCellProperties!.VerticalMerge!.Val!.Value);
            Assert.Equal(2, rows[2].Elements<W.TableCell>().ElementAt(1).TableCellProperties!.GridSpan!.Val!.Value);
            var tableGridStyle = document.MainDocumentPart.StyleDefinitionsPart!.Styles!
                .Elements<W.Style>().Single(style => style.StyleId == "TableGrid");
            Assert.Equal(W.StyleValues.Table, tableGridStyle.Type!.Value);
            Assert.Equal(6, tableGridStyle.StyleTableProperties!.TableBorders!.ChildElements.Count);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }

        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = exported.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var tableArtifact = Assert.Single(imported.Artifact.Document.Blocks).Table;
        Assert.Equal(3u, tableArtifact.GridColumns);
        Assert.Equal(2u, tableArtifact.Rows[0].RichCells[0].ColumnSpan);
        Assert.Equal(2u, tableArtifact.Rows[0].RichCells[0].RowSpan);
        Assert.Equal(DocumentTableVerticalMerge.Restart, tableArtifact.Rows[0].RichCells[0].VerticalMerge);
        Assert.Equal(DocumentTableVerticalMerge.Continue, tableArtifact.Rows[1].RichCells[0].VerticalMerge);
        Assert.False(tableArtifact.Rows[1].RichCells[0].Editable);
        Assert.NotNull(tableArtifact.Formatting);
        Assert.Equal(9000u, tableArtifact.Formatting.WidthDxa);
        Assert.Equal(240u, tableArtifact.Formatting.IndentDxa);
        Assert.Equal(new uint[] { 3000, 3000, 3000 }, tableArtifact.Formatting.ColumnWidthsDxa);
        Assert.Equal(80u, tableArtifact.Formatting.CellMarginsDxa.Top);
        Assert.Equal(120u, tableArtifact.Formatting.CellMarginsDxa.Start);
        Assert.Equal("445566", tableArtifact.Formatting.BorderColor);
        Assert.Equal(8u, tableArtifact.Formatting.BorderSize);
        Assert.Equal("E2E8F0", tableArtifact.Formatting.HeaderFill);
    }

    [Fact]
    public void DirectTableAuthoringRejectsIncompleteOrAmbiguousMergeGeometry()
    {
        var missingContinuation = MergedTableExportRequest();
        missingContinuation.Artifact.Document.Blocks[0].Table.Rows[0].RichCells[0].RowSpan = 3;
        var missingContinuationResponse = Invoke(missingContinuation);
        Assert.False(missingContinuationResponse.Ok);
        Assert.Equal("invalid_document_table", Assert.Single(missingContinuationResponse.Diagnostics).Code);

        var nonEmptyContinuation = MergedTableExportRequest();
        nonEmptyContinuation.Artifact.Document.Blocks[0].Table.Rows[1].Cells[0] = "ambiguous continuation text";
        var nonEmptyContinuationResponse = Invoke(nonEmptyContinuation);
        Assert.False(nonEmptyContinuationResponse.Ok);
        Assert.Equal("invalid_document_table", Assert.Single(nonEmptyContinuationResponse.Diagnostics).Code);

        var mismatchedSpan = MergedTableExportRequest();
        mismatchedSpan.Artifact.Document.Blocks[0].Table.Rows[1].RichCells[0].ColumnSpan = 1;
        var mismatchedSpanResponse = Invoke(mismatchedSpan);
        Assert.False(mismatchedSpanResponse.Ok);
        Assert.Equal("invalid_document_table", Assert.Single(mismatchedSpanResponse.Diagnostics).Code);

        var customStyle = MergedTableExportRequest();
        customStyle.Artifact.Document.Blocks[0].StyleId = "UnmodeledTableStyle";
        var customStyleResponse = Invoke(customStyle);
        Assert.False(customStyleResponse.Ok);
        Assert.Equal("unsupported_document_features", Assert.Single(customStyleResponse.Diagnostics).Code);

        var mismatchedWidths = MergedTableExportRequest();
        mismatchedWidths.Artifact.Document.Blocks[0].Table.Formatting.ColumnWidthsDxa[2] = 2999;
        var mismatchedWidthsResponse = Invoke(mismatchedWidths);
        Assert.False(mismatchedWidthsResponse.Ok);
        Assert.Equal("invalid_document_table", Assert.Single(mismatchedWidthsResponse.Diagnostics).Code);

        var invalidBorder = MergedTableExportRequest();
        invalidBorder.Artifact.Document.Blocks[0].Table.Formatting.BorderSize = 1;
        var invalidBorderResponse = Invoke(invalidBorder);
        Assert.False(invalidBorderResponse.Ok);
        Assert.Equal("invalid_document_table", Assert.Single(invalidBorderResponse.Diagnostics).Code);

        var noBorder = MergedTableExportRequest();
        noBorder.Artifact.Document.Blocks[0].Table.Formatting.BorderSize = 0;
        var noBorderResponse = Invoke(noBorder);
        Assert.True(noBorderResponse.Ok, Diagnostics(noBorderResponse));
        using (var stream = new MemoryStream(noBorderResponse.File.ToByteArray()))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            Assert.All(
                document.MainDocumentPart!.Document!.Body!.Descendants<W.TableBorders>().Single().ChildElements,
                border => Assert.Equal("nil", border.GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value));
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }
        var noBorderImported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = noBorderResponse.File,
        });
        Assert.Equal(0u, noBorderImported.Artifact.Document.Blocks[0].Table.Formatting.BorderSize);
    }

    [Fact]
    public void ImportedRecognizedDirectTableFormattingIsEditableAndPreservesResidual()
    {
        var authored = Invoke(MergedTableExportRequest());
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(AddRecognizedTableFormattingResidual(authored.File.ToByteArray())),
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var table = imported.Artifact.Document.Blocks[0].Table;
        Assert.NotNull(table.Formatting);
        table.Rows[0].Cells[0] = "Edited merged owner";
        table.Formatting.WidthDxa = 9600;
        table.Formatting.IndentDxa = 360;
        table.Formatting.ColumnWidthsDxa.Clear();
        table.Formatting.ColumnWidthsDxa.Add([2400, 3200, 4000]);
        table.Formatting.CellMarginsDxa = new DocumentTableCellMargins { Top = 40, Bottom = 60, Start = 80, End = 100 };
        table.Formatting.BorderColor = "AA3300";
        table.Formatting.BorderSize = 12;
        table.Formatting.HeaderFill = "FFF2CC";
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        using (var stream = new MemoryStream(edited.File.ToByteArray()))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var nativeTable = document.MainDocumentPart!.Document!.Body!.Elements<W.Table>().Single();
            Assert.Equal("9600", nativeTable.TableProperties!.TableWidth!.Width!.Value);
            Assert.Equal(360, nativeTable.TableProperties.TableIndentation!.Width!.Value);
            Assert.Equal(new[] { "2400", "3200", "4000" }, nativeTable.TableGrid!.Elements<W.GridColumn>().Select(column => column.Width!.Value));
            Assert.All(nativeTable.TableProperties.TableBorders!.ChildElements, border =>
            {
                Assert.Equal("AA3300", border.GetAttribute("color", "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value);
                Assert.Equal("12", border.GetAttribute("sz", "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value);
            });
            var rows = nativeTable.Elements<W.TableRow>().ToArray();
            Assert.Equal("Edited merged owner", rows[0].Elements<W.TableCell>().First().InnerText);
            Assert.Equal("FFF2CC", rows[0].Elements<W.TableCell>().First().TableCellProperties!.Shading!.Fill!.Value);
            Assert.Equal(480u, rows[1].TableRowProperties!.GetFirstChild<W.TableRowHeight>()!.Val!.Value);
            Assert.Equal(W.JustificationValues.Center, rows[2].Elements<W.TableCell>().First().Elements<W.Paragraph>().Single().ParagraphProperties!.Justification!.Val!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }

        var roundTrip = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = edited.File,
        });
        Assert.True(roundTrip.Ok, Diagnostics(roundTrip));
        var roundTripFormatting = roundTrip.Artifact.Document.Blocks[0].Table.Formatting;
        Assert.Equal(9600u, roundTripFormatting.WidthDxa);
        Assert.Equal(360u, roundTripFormatting.IndentDxa);
        Assert.Equal(new uint[] { 2400, 3200, 4000 }, roundTripFormatting.ColumnWidthsDxa);
        Assert.Equal(40u, roundTripFormatting.CellMarginsDxa.Top);
        Assert.Equal(100u, roundTripFormatting.CellMarginsDxa.End);
        Assert.Equal("AA3300", roundTripFormatting.BorderColor);
        Assert.Equal(12u, roundTripFormatting.BorderSize);
        Assert.Equal("FFF2CC", roundTripFormatting.HeaderFill);

        var removed = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        removed.Artifact.Document.Blocks[0].Table.Formatting = null;
        var removedRejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = removed.Artifact,
        });
        Assert.False(removedRejected.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(removedRejected.Diagnostics).Code);

        var unformattedRequest = MergedTableExportRequest();
        unformattedRequest.Artifact.Document.Blocks[0].Table.Formatting = null;
        var unformatted = Invoke(unformattedRequest);
        Assert.True(unformatted.Ok, Diagnostics(unformatted));
        var unrecognized = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = unformatted.File,
        });
        Assert.Null(unrecognized.Artifact.Document.Blocks[0].Table.Formatting);
        unrecognized.Artifact.Document.Blocks[0].Table.Formatting = MergedTableExportRequest().Artifact.Document.Blocks[0].Table.Formatting.Clone();
        var unrecognizedRejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = unrecognized.Artifact,
        });
        Assert.False(unrecognizedRejected.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(unrecognizedRejected.Diagnostics).Code);
    }

    [Fact]
    public void DirectAuthoringBuildsValidatedSharedMultilevelNumberingGraph()
    {
        var exported = Invoke(DirectNumberingExportRequest());
        Assert.True(exported.Ok, Diagnostics(exported));
        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var mainPart = document.MainDocumentPart!;
            var numbering = mainPart.NumberingDefinitionsPart!.Numbering!;
            var abstractNumbering = Assert.Single(numbering.Elements<W.AbstractNum>());
            Assert.Equal(9, abstractNumbering.AbstractNumberId!.Value);
            var levels = abstractNumbering.Elements<W.Level>().ToArray();
            Assert.Equal([0, 2], levels.Select(level => level.LevelIndex!.Value).ToArray());
            Assert.Equal(W.NumberFormatValues.UpperLetter, levels[0].NumberingFormat!.Val!.Value);
            Assert.Equal(3, levels[0].StartNumberingValue!.Val!.Value);
            Assert.Equal("%1)", levels[0].LevelText!.Val!.Value);
            Assert.Equal(W.NumberFormatValues.LowerRoman, levels[1].NumberingFormat!.Val!.Value);
            Assert.Equal(5, levels[1].StartNumberingValue!.Val!.Value);
            Assert.Equal("%1.%2.%3.", levels[1].LevelText!.Val!.Value);
            var instances = numbering.Elements<W.NumberingInstance>().ToArray();
            Assert.Equal([77, 78], instances.Select(instance => instance.NumberID!.Value).ToArray());
            Assert.All(instances, instance => Assert.Equal(9, instance.AbstractNumId!.Val!.Value));
            var paragraphs = mainPart.Document!.Body!.Elements<W.Paragraph>().ToArray();
            Assert.Equal([77, 77, 78], paragraphs.Select(paragraph => paragraph.ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value).ToArray());
            Assert.Equal([0, 2, 0], paragraphs.Select(paragraph => paragraph.ParagraphProperties!.NumberingProperties!.NumberingLevelReference!.Val!.Value).ToArray());
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }

        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = exported.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        Assert.Collection(imported.Artifact.Document.Blocks,
            first =>
            {
                Assert.Equal(77u, first.Paragraph.Numbering.NumberingId);
                Assert.Equal(9u, first.Paragraph.Numbering.AbstractNumberingId);
                Assert.Equal("upperLetter", first.Paragraph.Numbering.NumberFormat);
            },
            second =>
            {
                Assert.Equal(2u, second.Paragraph.Numbering.Level);
                Assert.Equal("lowerRoman", second.Paragraph.Numbering.NumberFormat);
                Assert.Equal(5u, second.Paragraph.Numbering.Start);
            },
            third => Assert.Equal(78u, third.Paragraph.Numbering.NumberingId));
    }

    [Fact]
    public void DirectNumberingAuthoringRejectsConflictingOrUnmodeledGraphs()
    {
        var conflictingLevel = DirectNumberingExportRequest();
        conflictingLevel.Artifact.Document.Blocks[1].Paragraph.Numbering.Level = 0;
        var conflictingLevelResponse = Invoke(conflictingLevel);
        Assert.False(conflictingLevelResponse.Ok);
        Assert.Equal("invalid_document_numbering", Assert.Single(conflictingLevelResponse.Diagnostics).Code);

        var conflictingInstance = DirectNumberingExportRequest();
        conflictingInstance.Artifact.Document.Blocks[1].Paragraph.Numbering.AbstractNumberingId = 10;
        var conflictingInstanceResponse = Invoke(conflictingInstance);
        Assert.False(conflictingInstanceResponse.Ok);
        Assert.Equal("invalid_document_numbering", Assert.Single(conflictingInstanceResponse.Diagnostics).Code);

        var linkedStyle = DirectNumberingExportRequest();
        linkedStyle.Artifact.Document.Blocks[0].Paragraph.Numbering.NumberingStyleId = "AgentNumbering";
        var linkedStyleResponse = Invoke(linkedStyle);
        Assert.False(linkedStyleResponse.Ok);
        Assert.Equal("unsupported_document_features", Assert.Single(linkedStyleResponse.Diagnostics).Code);

        var zeroStart = DirectNumberingExportRequest();
        zeroStart.Artifact.Document.Blocks[0].Paragraph.Numbering.Start = 0;
        var zeroStartResponse = Invoke(zeroStart);
        Assert.False(zeroStartResponse.Ok);
        Assert.Equal("invalid_document_numbering", Assert.Single(zeroStartResponse.Diagnostics).Code);
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
    public void SourcePreservingExportEditsCompleteDirectNumberingGroupThroughInstanceOverride()
    {
        var authored = Invoke(ExportRequest(includeSecondParagraph: true));
        var source = AddDirectNumbering(authored.File.ToByteArray(), numberAllParagraphs: true, level: 2);
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(source),
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        Assert.All(imported.Artifact.Document.Blocks, block =>
        {
            Assert.True(block.Source.Editable);
            Assert.Equal(77u, block.Paragraph.Numbering.NumberingId);
            Assert.Equal(2u, block.Paragraph.Numbering.Level);
            block.Paragraph.Numbering.NumberFormat = "lowerRoman";
            block.Paragraph.Numbering.Start = 5;
            block.Paragraph.Numbering.LevelText = "%1.%2.%3.";
        });
        imported.Artifact.Document.Blocks[0].Paragraph.Text = "Edited grouped list title";
        imported.Artifact.Document.Blocks[0].Paragraph.Runs[0].Text = "Edited grouped list title";

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
            var numbering = document.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
            var abstractLevel = numbering.Elements<W.AbstractNum>().Single().Elements<W.Level>().Single();
            Assert.Equal(W.NumberFormatValues.UpperLetter, abstractLevel.NumberingFormat!.Val!.Value);
            Assert.Equal(3, abstractLevel.StartNumberingValue!.Val!.Value);
            Assert.Equal("%1)", abstractLevel.LevelText!.Val!.Value);
            var levelOverride = numbering.Elements<W.NumberingInstance>().Single().Elements<W.LevelOverride>().Single();
            var localLevel = levelOverride.Elements<W.Level>().Single();
            Assert.Equal(2, levelOverride.LevelIndex!.Value);
            Assert.Equal(W.NumberFormatValues.LowerRoman, localLevel.NumberingFormat!.Val!.Value);
            Assert.Equal(5, localLevel.StartNumberingValue!.Val!.Value);
            Assert.Equal("%1.%2.%3.", localLevel.LevelText!.Val!.Value);
            Assert.Equal("Edited grouped list title", document.MainDocumentPart.Document!.Body!.Elements<W.Paragraph>().First().InnerText);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }

        var roundTrip = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = exported.File,
        });
        Assert.True(roundTrip.Ok, Diagnostics(roundTrip));
        Assert.All(roundTrip.Artifact.Document.Blocks, block =>
        {
            Assert.Equal("lowerRoman", block.Paragraph.Numbering.NumberFormat);
            Assert.Equal(5u, block.Paragraph.Numbering.Start);
            Assert.Equal("%1.%2.%3.", block.Paragraph.Numbering.LevelText);
        });
    }

    [Fact]
    public void NumberingDefinitionEditRejectsPartialGroupAndStyleInheritedGraph()
    {
        var authored = Invoke(ExportRequest(includeSecondParagraph: true));
        var grouped = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(AddDirectNumbering(authored.File.ToByteArray(), numberAllParagraphs: true)),
        });
        grouped.Artifact.Document.Blocks[0].Paragraph.Numbering.Start = 11;
        var partial = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = grouped.Artifact,
        });
        Assert.False(partial.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(partial.Diagnostics).Code);

        var inherited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(AddStyleInheritedNumbering(authored.File.ToByteArray())),
        });
        inherited.Artifact.Document.Blocks[1].Paragraph.Numbering.Start = 12;
        var styleLinked = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = inherited.Artifact,
        });
        Assert.False(styleLinked.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(styleLinked.Diagnostics).Code);

        var crossPart = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(AddUnusedStyleNumberingReference(AddDirectNumbering(authored.File.ToByteArray()))),
        });
        crossPart.Artifact.Document.Blocks[1].Paragraph.Numbering.Start = 13;
        var crossPartRejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = crossPart.Artifact,
        });
        Assert.False(crossPartRejected.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(crossPartRejected.Diagnostics).Code);
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
        Assert.All(imported.Artifact.OpaqueOpc.Parts, part => Assert.False(string.IsNullOrWhiteSpace(part.ContentType)));
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
    public void SourcePreservingExportPatchesPlainTextInsideReadOnlyParagraph()
    {
        var authored = Invoke(ExportRequest(includeSecondParagraph: true));
        var source = AddUnsupportedParagraphProperty(authored.File.ToByteArray());
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
        Assert.True(block.Source.TextPatchable);
        block.TextPatches.Add(new DocumentTextPatch
        {
            Search = "Editable",
            Replacement = "Reviewed",
            SourceTextSha256 = Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes("Editable paragraph"))).ToLowerInvariant(),
        });

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
            Assert.Equal("Reviewed paragraph", paragraph.InnerText);
            Assert.NotNull(paragraph.ParagraphProperties?.GetFirstChild<W.WidowControl>());
        }
        var roundTrip = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = exported.File,
        });
        Assert.True(roundTrip.Ok, Diagnostics(roundTrip));
        Assert.Equal("Reviewed paragraph", roundTrip.Artifact.Document.Blocks[1].Paragraph.Text);
        Assert.False(roundTrip.Artifact.Document.Blocks[1].Source.Editable);
        Assert.True(roundTrip.Artifact.Document.Blocks[1].Source.TextPatchable);

        var tampered = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(source),
        });
        tampered.Artifact.Document.Blocks[1].Source.TextPatchable = false;
        var tamperedRejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = tampered.Artifact,
        });
        Assert.False(tamperedRejected.Ok);
        Assert.Equal("document_source_binding_mismatch", Assert.Single(tamperedRejected.Diagnostics).Code);

        var invalidBlock = block.Clone();
        invalidBlock.TextPatches.Clear();
        invalidBlock.TextPatches.Add(new DocumentTextPatch
        {
            Search = "Editable",
            Replacement = "\ud800",
            SourceTextSha256 = Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes("Editable paragraph"))).ToLowerInvariant(),
        });
        var invalidBlockException = Assert.Throws<CodecException>(() => DocxPlainTextPatchCodec.Validate(invalidBlock));
        Assert.Equal("invalid_document_text_patch", invalidBlockException.Code);
    }

    [Fact]
    public void SourcePreservingExportPatchesAdjacentSameFormatRunSpans()
    {
        var paragraphAuthored = Invoke(ExportRequest(includeSecondParagraph: true));
        var paragraphSource = AddFragmentedReadOnlyParagraph(paragraphAuthored.File.ToByteArray());
        var paragraphImported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(paragraphSource),
        });
        Assert.True(paragraphImported.Ok, Diagnostics(paragraphImported));
        var paragraphBlock = paragraphImported.Artifact.Document.Blocks[1];
        Assert.False(paragraphBlock.Source.Editable);
        Assert.True(paragraphBlock.Source.TextPatchable);
        paragraphBlock.TextPatches.Add(new DocumentTextPatch
        {
            Search = "Editable",
            Replacement = "Reviewed",
            SourceTextSha256 = HashText("Editable paragraph"),
        });

        var paragraphExported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = paragraphImported.Artifact,
        });
        Assert.True(paragraphExported.Ok, Diagnostics(paragraphExported));
        using (var stream = new MemoryStream(paragraphExported.File.ToByteArray()))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var paragraph = document.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().ElementAt(1);
            Assert.Equal("Reviewed paragraph", paragraph.InnerText);
            Assert.Equal(new[] { "Reviewed", " paragraph" }, paragraph.Elements<W.Run>().Select(run => run.InnerText));
            Assert.NotNull(paragraph.ParagraphProperties?.GetFirstChild<W.WidowControl>());
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }
        var paragraphRoundTrip = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = paragraphExported.File,
        });
        Assert.True(paragraphRoundTrip.Ok, Diagnostics(paragraphRoundTrip));
        Assert.Equal("Reviewed paragraph", paragraphRoundTrip.Artifact.Document.Blocks[1].Paragraph.Text);
        Assert.True(paragraphRoundTrip.Artifact.Document.Blocks[1].Source.TextPatchable);

        var tableAuthored = Invoke(ExportRequest());
        var tableSource = AddFragmentedTableCell(tableAuthored.File.ToByteArray());
        var tableImported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(tableSource),
        });
        Assert.True(tableImported.Ok, Diagnostics(tableImported));
        var table = tableImported.Artifact.Document.Blocks[1].Table;
        Assert.False(table.Rows[0].RichCells[0].Editable);
        Assert.True(table.Rows[0].RichCells[0].TextPatchable);
        table.TextPatches.Add(new DocumentTableTextPatch
        {
            Row = 0,
            Column = 0,
            Search = "Revenue",
            Replacement = "Net revenue",
            SourceTextSha256 = HashText("Revenue"),
        });
        var tableExported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = tableImported.Artifact,
        });
        Assert.True(tableExported.Ok, Diagnostics(tableExported));
        using (var stream = new MemoryStream(tableExported.File.ToByteArray()))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var cell = document.MainDocumentPart!.Document!.Body!.Elements<W.Table>().Single().Elements<W.TableRow>().Single().Elements<W.TableCell>().First();
            Assert.Equal("Net revenue", cell.InnerText);
            Assert.Equal(new[] { "Net revenue", string.Empty }, cell.Elements<W.Paragraph>().Single().Elements<W.Run>().Select(run => run.InnerText));
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }
        var tableRoundTrip = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = tableExported.File,
        });
        Assert.True(tableRoundTrip.Ok, Diagnostics(tableRoundTrip));
        Assert.Equal("Net revenue", tableRoundTrip.Artifact.Document.Blocks[1].Table.Rows[0].Cells[0]);
        Assert.True(tableRoundTrip.Artifact.Document.Blocks[1].Table.Rows[0].RichCells[0].TextPatchable);
    }

    [Theory]
    [InlineData(true, false, "different formatting")]
    [InlineData(false, true, "empty-run gap")]
    public void SourcePreservingRunSpanPatchRejectsFormattingAndEmptyRunGaps(
        bool mixedFormatting,
        bool emptyRunGap,
        string expectedReason)
    {
        var authored = Invoke(ExportRequest(includeSecondParagraph: true));
        var source = AddFragmentedReadOnlyParagraph(authored.File.ToByteArray(), mixedFormatting, emptyRunGap);
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(source),
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var block = imported.Artifact.Document.Blocks[1];
        Assert.True(block.Source.TextPatchable);
        block.TextPatches.Add(new DocumentTextPatch
        {
            Search = "Editable",
            Replacement = "Reviewed",
            SourceTextSha256 = HashText("Editable paragraph"),
        });
        var rejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.False(rejected.Ok);
        var diagnostic = Assert.Single(rejected.Diagnostics);
        Assert.Equal("unsupported_document_edit", diagnostic.Code);
        Assert.Contains(expectedReason, diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TableTextPatchDoesNotEnterControlsOrCrossParagraphs()
    {
        var authored = Invoke(ExportRequest());
        var controlled = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(WrapFirstTableCellRunInContentControl(authored.File.ToByteArray())),
        });
        Assert.True(controlled.Ok, Diagnostics(controlled));
        Assert.False(controlled.Artifact.Document.Blocks[1].Table.Rows[0].RichCells[0].TextPatchable);

        var crossParagraph = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(AddSecondTableParagraph(authored.File.ToByteArray())),
        });
        Assert.True(crossParagraph.Ok, Diagnostics(crossParagraph));
        crossParagraph.Artifact.Document.Blocks[1].Table.TextPatches.Add(new DocumentTableTextPatch
        {
            Row = 0,
            Column = 0,
            Search = "Revenue detail",
            Replacement = "Net revenue detail",
            SourceTextSha256 = HashText("Revenue detail"),
        });
        var rejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = crossParagraph.Artifact,
        });
        Assert.False(rejected.Ok);
        var diagnostic = Assert.Single(rejected.Diagnostics);
        Assert.Equal("unsupported_document_edit", diagnostic.Code);
        Assert.Contains("paragraph boundary", diagnostic.Message, StringComparison.Ordinal);
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
    public void SourcePreservingExportRetainsPreExistingOpaqueValidationWarning()
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
        Assert.False(imported.Artifact.Document.Blocks[0].Source.Editable);
        var response = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(response.Ok, Diagnostics(response));
        Assert.Contains(response.Diagnostics, item => item.Code == "source_openxml_validation_warnings_preserved");
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
    public void WholeParagraphBookmarkAuthorsImportsAndFailsClosedOnMutation()
    {
        var exported = Invoke(BookmarkExportRequest());
        Assert.True(exported.Ok, Diagnostics(exported));
        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var package = WordprocessingDocument.Open(stream, false))
        {
            var body = package.MainDocumentPart!.Document!.Body!;
            var start = Assert.Single(body.Descendants<W.BookmarkStart>());
            var end = Assert.Single(body.Descendants<W.BookmarkEnd>());
            Assert.Equal("TargetBookmark", start.Name?.Value);
            Assert.Equal(start.Id?.Value, end.Id?.Value);
            Assert.Equal("TargetBookmark", Assert.Single(body.Descendants<W.Hyperlink>()).Anchor?.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = exported.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var bookmark = Assert.Single(imported.Artifact.Document.Bookmarks);
        Assert.Equal("TargetBookmark", bookmark.Name);
        Assert.Equal("document/block/2", bookmark.TargetBlockId);
        Assert.Equal(bookmark.TargetBlockId, bookmark.EndTargetBlockId);
        Assert.Equal("0", bookmark.NativeId);
        Assert.False(bookmark.Source.Editable);

        var unchanged = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(unchanged.Ok, Diagnostics(unchanged));

        bookmark.Name = "RenamedBookmark";
        var rejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_document_bookmark_edit", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void PlainTextFootnotesAndEndnotesAuthorImportEditAndFailClosed()
    {
        var authored = Invoke(NoteExportRequest());
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var package = WordprocessingDocument.Open(stream, false))
        {
            var owner = package.MainDocumentPart!;
            Assert.Single(owner.Document!.Body!.Descendants<W.FootnoteReference>());
            Assert.Single(owner.Document.Body.Descendants<W.EndnoteReference>());
            Assert.Equal("Source-free footnote", owner.FootnotesPart!.Footnotes!
                .Elements<W.Footnote>().Single(note => note.Id?.Value == 1).InnerText.Trim());
            Assert.Equal("Source-free endnote", owner.EndnotesPart!.Endnotes!
                .Elements<W.Endnote>().Single(note => note.Id?.Value == 1).InnerText.Trim());
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        Assert.Collection(imported.Artifact.Document.Notes,
            note =>
            {
                Assert.Equal(DocumentNoteKind.Footnote, note.Kind);
                Assert.Equal("Source-free footnote", note.Text);
                Assert.Equal("1", note.NativeId);
                Assert.True(note.Source.Editable);
                Assert.Equal("word/footnotes.xml", note.Source.PartPath);
                Assert.NotEmpty(note.Source.RelationshipId);
                Assert.NotEmpty(note.Source.NoteElementSha256);
                Assert.NotEmpty(note.Source.AnchorSha256);
            },
            note =>
            {
                Assert.Equal(DocumentNoteKind.Endnote, note.Kind);
                Assert.Equal("Source-free endnote", note.Text);
                Assert.Equal("1", note.NativeId);
                Assert.True(note.Source.Editable);
                Assert.Equal("word/endnotes.xml", note.Source.PartPath);
            });
        Assert.All(imported.Artifact.Document.Blocks, block => Assert.False(block.Source.Editable));

        var unchanged = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact.Clone(),
        });
        Assert.True(unchanged.Ok, Diagnostics(unchanged));
        Assert.Equal(authored.File.ToByteArray(), unchanged.File.ToByteArray());

        var editedArtifact = imported.Artifact.Clone();
        editedArtifact.Document.Notes[0].Text = "Edited footnote";
        editedArtifact.Document.Notes[1].Text = "Edited endnote";
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = editedArtifact,
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        var reimported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = edited.File,
        });
        Assert.True(reimported.Ok, Diagnostics(reimported));
        Assert.Equal(["Edited footnote", "Edited endnote"], reimported.Artifact.Document.Notes.Select(note => note.Text));

        var moved = imported.Artifact.Clone();
        moved.Document.Notes[0].TargetBlockId = moved.Document.Notes[1].TargetBlockId;
        var rejectedMove = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = moved,
        });
        Assert.False(rejectedMove.Ok);
        Assert.Equal("invalid_document_note", Assert.Single(rejectedMove.Diagnostics).Code);

        var removed = imported.Artifact.Clone();
        removed.Document.Notes.RemoveAt(1);
        var rejectedTopology = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = removed,
        });
        Assert.False(rejectedTopology.Ok);
        Assert.Equal("document_note_topology_changed", Assert.Single(rejectedTopology.Diagnostics).Code);
    }

    [Fact]
    public void RichFootnoteBodyRemainsOpaqueAndBytePreserved()
    {
        var authored = Invoke(NoteExportRequest());
        Assert.True(authored.Ok, Diagnostics(authored));
        var richSource = AddSecondFootnoteParagraph(authored.File.ToByteArray());
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(richSource),
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var endnote = Assert.Single(imported.Artifact.Document.Notes);
        Assert.Equal(DocumentNoteKind.Endnote, endnote.Kind);
        Assert.False(imported.Artifact.Document.Blocks[0].Source.Editable);

        var unchanged = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact.Clone(),
        });
        Assert.True(unchanged.Ok, Diagnostics(unchanged));
        Assert.Equal(richSource, unchanged.File.ToByteArray());

        imported.Artifact.Document.Blocks[0].Paragraph.Text = "Attempted rich-note anchor edit";
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
    public void BibliographyCitationSliceAuthorsImportsEditsAndFailsClosedForIrregularGraphs()
    {
        var authored = Invoke(BibliographyExportRequest());
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var package = WordprocessingDocument.Open(stream, false))
        {
            var mainPart = package.MainDocumentPart!;
            var part = Assert.Single(mainPart.CustomXmlParts);
            var sources = new B.Sources();
            sources.Load(part);
            Assert.Contains("AgentSource", sources.OuterXml);
            var field = Assert.Single(mainPart.Document!.Descendants<W.SimpleField>());
            Assert.Equal(" CITATION AgentSource ", field.Instruction?.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var bibliography = Assert.IsType<DocumentBibliography>(imported.Artifact.Document.Bibliography);
        Assert.True(bibliography.Source.Editable);
        Assert.NotEmpty(bibliography.Source.PartSha256);
        Assert.Equal("AgentSource", Assert.Single(bibliography.Sources).Tag);
        var citation = Assert.Single(imported.Artifact.Document.Blocks);
        Assert.Equal(DocumentBlock.ContentOneofCase.Citation, citation.ContentCase);
        Assert.True(citation.Source.Editable);
        Assert.NotEmpty(citation.Source.ResidualSha256);
        Assert.Equal("(Lovelace, 1843)", citation.Citation.Display);
        Assert.Single(imported.Artifact.Document.Bookmarks);

        bibliography.Sources[0].Fields["title"] = "Notes on the Analytical Engine";
        bibliography.Sources[0].Authors[0].First = "Augusta Ada";
        citation.Citation.Display = "(Lovelace, 1843, revised)";
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        var roundTrip = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = edited.File,
        });
        Assert.True(roundTrip.Ok, Diagnostics(roundTrip));
        Assert.Equal("Notes on the Analytical Engine", roundTrip.Artifact.Document.Bibliography.Sources[0].Fields["title"]);
        Assert.Equal("Augusta Ada", roundTrip.Artifact.Document.Bibliography.Sources[0].Authors[0].First);
        Assert.Equal("(Lovelace, 1843, revised)", roundTrip.Artifact.Document.Blocks[0].Citation.Display);

        roundTrip.Artifact.Document.Bibliography.Sources[0].Tag = "RenamedSource";
        roundTrip.Artifact.Document.Blocks[0].Citation.Tag = "RenamedSource";
        var renamed = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = roundTrip.Artifact,
        });
        Assert.False(renamed.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(renamed.Diagnostics).Code);

        var irregularBytes = MakeBibliographyContributorRoleIrregular(authored.File.ToByteArray());
        var irregular = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(irregularBytes),
        });
        Assert.True(irregular.Ok, Diagnostics(irregular));
        Assert.Null(irregular.Artifact.Document.Bibliography);
        Assert.Equal(DocumentBlock.ContentOneofCase.Paragraph, irregular.Artifact.Document.Blocks[0].ContentCase);
        Assert.False(irregular.Artifact.Document.Blocks[0].Source.Editable);
        var unchanged = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = irregular.Artifact,
        });
        Assert.True(unchanged.Ok, Diagnostics(unchanged));
        irregular.Artifact.Document.Blocks[0].Paragraph.Text = "Unsafe reconstruction";
        var rejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = irregular.Artifact,
        });
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_document_edit", Assert.Single(rejected.Diagnostics).Code);
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

    [Fact]
    public void ComplexTocFieldAuthorsImportsEditsAndCarriesExplicitRefreshHint()
    {
        var authored = Invoke(TocExportRequest());
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var settings = document.MainDocumentPart!.DocumentSettingsPart!.Settings!;
            Assert.True(settings.GetFirstChild<W.UpdateFieldsOnOpen>()?.Val?.Value);
            var paragraph = document.MainDocumentPart.Document!.Body!.Elements<W.Paragraph>().Single();
            var run = paragraph.Elements<W.Run>().Single();
            Assert.Equal(W.FieldCharValues.Begin, ((W.FieldChar)run.ChildElements[0]).FieldCharType?.Value);
            Assert.Equal(" TOC \\o \"1-3\" \\h \\z \\u ", run.Elements<W.FieldCode>().Single().Text);
            Assert.Equal("Refresh in Word", run.Elements<W.Text>().Single().Text);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }

        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        Assert.True(imported.Artifact.Document.UpdateFields);
        var block = Assert.Single(imported.Artifact.Document.Blocks);
        Assert.Equal(DocumentBlock.ContentOneofCase.Field, block.ContentCase);
        Assert.True(block.Field.Complex);
        Assert.True(block.Source.Editable);
        Assert.Equal("TOC \\o \"1-3\" \\h \\z \\u", block.Field.Instruction);

        block.Field.Instruction = "TOC \\o \"1-4\" \\h \\z \\u";
        block.Field.Display = "Update before delivery";
        imported.Artifact.Document.UpdateFields = false;
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        using (var stream = new MemoryStream(edited.File.ToByteArray()))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            Assert.Null(document.MainDocumentPart!.DocumentSettingsPart!.Settings!.GetFirstChild<W.UpdateFieldsOnOpen>());
            Assert.Equal(" TOC \\o \"1-4\" \\h \\z \\u ", document.MainDocumentPart.Document!.Descendants<W.FieldCode>().Single().Text);
            Assert.Equal("Update before delivery", document.MainDocumentPart.Document.Descendants<W.Text>().Single().Text);
        }
        var secondImport = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = edited.File,
        });
        Assert.True(secondImport.Ok, Diagnostics(secondImport));
        Assert.False(secondImport.Artifact.Document.UpdateFields);
        Assert.Equal("TOC \\o \"1-4\" \\h \\z \\u", secondImport.Artifact.Document.Blocks[0].Field.Instruction);

        var withoutSettings = Invoke(FieldExportRequest());
        var plainImport = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = withoutSettings.File,
        });
        using (var plainStream = new MemoryStream(withoutSettings.File.ToByteArray()))
        using (var plainDocument = WordprocessingDocument.Open(plainStream, false))
            Assert.Null(plainDocument.MainDocumentPart!.DocumentSettingsPart);
        plainImport.Artifact.Document.UpdateFields = true;
        var added = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = plainImport.Artifact,
        });
        Assert.True(added.Ok, Diagnostics(added));
        using var addedStream = new MemoryStream(added.File.ToByteArray());
        using var addedDocument = WordprocessingDocument.Open(addedStream, false);
        Assert.True(addedDocument.MainDocumentPart!.DocumentSettingsPart!.Settings!.GetFirstChild<W.UpdateFieldsOnOpen>()?.Val?.Value);
    }

    [Fact]
    public void RefreshedCrossParagraphTocGraphIsOpaqueAndReadOnlyAsOneFieldSpan()
    {
        var authored = Invoke(TocExportRequest());
        var source = ExpandTocFieldAcrossParagraphs(authored.File.ToByteArray());
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(source),
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        Assert.Equal(3, imported.Artifact.Document.Blocks.Count);
        Assert.All(imported.Artifact.Document.Blocks, block =>
        {
            Assert.Equal(DocumentBlock.ContentOneofCase.Opaque, block.ContentCase);
            Assert.False(block.Source.Editable);
        });

        var unchanged = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(unchanged.Ok, Diagnostics(unchanged));

        imported.Artifact.Document.Blocks[1].Opaque.Text = "Unsafe cached-result edit";
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
    public void TrackedChangeSliceAuthorsImportsAndValidatesWholeParagraphInsertionsAndDeletions()
    {
        var authored = Invoke(TrackedChangeExportRequest());
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var body = document.MainDocumentPart!.Document!.Body!;
            var insertion = body.Descendants<W.InsertedRun>().Single();
            var deletion = body.Descendants<W.DeletedRun>().Single();
            Assert.Equal("1", insertion.Id?.Value);
            Assert.Equal("2", deletion.Id?.Value);
            Assert.Equal("Reviewer", insertion.Author?.Value);
            Assert.Equal("Removed wording", deletion.Descendants<W.DeletedText>().Single().Text);
            Assert.Empty(deletion.Descendants<W.Text>());
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }

        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        Assert.Equal(DocumentBlock.ContentOneofCase.Change, imported.Artifact.Document.Blocks[0].ContentCase);
        Assert.Equal(DocumentChangeType.Insert, imported.Artifact.Document.Blocks[0].Change.Type);
        Assert.Equal("Added wording", imported.Artifact.Document.Blocks[0].Change.Text);
        Assert.Equal("1", imported.Artifact.Document.Blocks[0].Source.NativeRevisionId);
        Assert.True(imported.Artifact.Document.Blocks[0].Source.Editable);
        Assert.Equal(DocumentChangeType.Delete, imported.Artifact.Document.Blocks[1].Change.Type);
        Assert.Equal("Removed wording", imported.Artifact.Document.Blocks[1].Change.Text);
        Assert.Equal("2", imported.Artifact.Document.Blocks[1].Source.NativeRevisionId);
    }

    [Fact]
    public void SourcePreservingExportEditsTrackedChangeSemanticsAndKeepsRevisionIdentityAndFormatting()
    {
        var authored = Invoke(TrackedChangeExportRequest(includeDeletion: false));
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = AddTrackedChangeFormatting(authored.File.ToByteArray());
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(source),
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var block = imported.Artifact.Document.Blocks.Single();
        Assert.Equal(DocumentBlock.ContentOneofCase.Change, block.ContentCase);
        Assert.True(block.Source.Editable);
        Assert.Equal("1", block.Source.NativeRevisionId);
        Assert.NotEmpty(block.Source.ResidualSha256);

        block.Change.Text = "Edited insertion";
        block.Change.Author = "Lead reviewer";
        block.Change.Date = "2026-07-17T08:30:00Z";
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
            var insertion = document.MainDocumentPart!.Document!.Descendants<W.InsertedRun>().Single();
            Assert.Equal("1", insertion.Id?.Value);
            Assert.Equal("Lead reviewer", insertion.Author?.Value);
            Assert.Equal("Edited insertion", insertion.InnerText);
            Assert.NotNull(insertion.Descendants<W.Bold>().SingleOrDefault());
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }

        block.Change.Type = DocumentChangeType.Delete;
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
    public void IrregularTrackedChangeTopologyIsVisibleButReadOnlyAndPreservedUnchanged()
    {
        var authored = Invoke(TrackedChangeExportRequest(includeDeletion: false));
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = AddSecondTrackedChangeRun(authored.File.ToByteArray());
        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = ByteString.CopyFrom(source),
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        var block = Assert.Single(imported.Artifact.Document.Blocks);
        Assert.Equal(DocumentBlock.ContentOneofCase.Change, block.ContentCase);
        Assert.Equal("Added wording continuation", block.Change.Text);
        Assert.False(block.Source.Editable);
        Assert.Empty(block.Source.ResidualSha256);

        var unchanged = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(unchanged.Ok, Diagnostics(unchanged));
        Assert.Equal(source, unchanged.File.ToByteArray());

        block.Change.Text = "Attempted edit";
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
    public void TrackRevisionsSettingAuthorsImportsAndSourceBoundTogglesWithoutChangingRevisionMarkup()
    {
        var request = TrackedChangeExportRequest(trackRevisions: true);
        var authored = Invoke(request);
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var package = WordprocessingDocument.Open(stream, false))
            Assert.NotNull(package.MainDocumentPart!.DocumentSettingsPart!.Settings!.GetFirstChild<W.TrackRevisions>());

        var imported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = authored.File,
        });
        Assert.True(imported.Ok, Diagnostics(imported));
        Assert.True(imported.Artifact.Document.TrackRevisions);
        imported.Artifact.Document.TrackRevisions = false;

        var exported = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = imported.Artifact,
        });
        Assert.True(exported.Ok, Diagnostics(exported));
        using var outputStream = new MemoryStream(exported.File.ToByteArray());
        using var output = WordprocessingDocument.Open(outputStream, false);
        Assert.Null(output.MainDocumentPart!.DocumentSettingsPart!.Settings!.GetFirstChild<W.TrackRevisions>());
        Assert.Single(output.MainDocumentPart.Document!.Descendants<W.InsertedRun>());
        Assert.Single(output.MainDocumentPart.Document.Descendants<W.DeletedRun>());
    }

    [Fact]
    public void RevisionFinalizationAcceptsAndRejectsBoundedChangesWithExactPartAudit()
    {
        var authored = Invoke(TrackedChangeExportRequest(trackRevisions: true));
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = AddAllTrackedChangeFormatting(authored.File.ToByteArray());
        var sourceHash = Convert.ToHexString(SHA256.HashData(source)).ToLowerInvariant();

        var accepted = Invoke(FinalizeRevisionRequest(source, sourceHash, DocumentRevisionFinalizationMode.Accept));
        Assert.True(accepted.Ok, Diagnostics(accepted));
        Assert.Equal(DocumentRevisionFinalizationMode.Accept, accepted.RevisionFinalization.Mode);
        Assert.Equal(sourceHash, accepted.RevisionFinalization.SourceSha256);
        Assert.Equal(1U, accepted.RevisionFinalization.InsertionCount);
        Assert.Equal(1U, accepted.RevisionFinalization.DeletionCount);
        Assert.True(accepted.RevisionFinalization.TrackingBefore);
        Assert.False(accepted.RevisionFinalization.TrackingAfter);
        Assert.Equal(new[] { "word/document.xml", "word/settings.xml" }, accepted.RevisionFinalization.ChangedParts);
        Assert.Equal(
            Convert.ToHexString(SHA256.HashData(accepted.File.Span)).ToLowerInvariant(),
            accepted.RevisionFinalization.OutputSha256);
        using (var stream = new MemoryStream(accepted.File.ToByteArray()))
        using (var package = WordprocessingDocument.Open(stream, false))
        {
            var paragraphs = package.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().ToArray();
            Assert.Equal("Added wording", paragraphs[0].InnerText);
            Assert.NotNull(paragraphs[0].Descendants<W.Bold>().SingleOrDefault());
            Assert.Equal(string.Empty, paragraphs[1].InnerText);
            Assert.Empty(package.MainDocumentPart.Document.Descendants<W.InsertedRun>());
            Assert.Empty(package.MainDocumentPart.Document.Descendants<W.DeletedRun>());
            Assert.Null(package.MainDocumentPart.DocumentSettingsPart!.Settings!.GetFirstChild<W.TrackRevisions>());
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var rejected = Invoke(FinalizeRevisionRequest(
            source,
            sourceHash,
            DocumentRevisionFinalizationMode.Reject,
            keepTracking: true));
        Assert.True(rejected.Ok, Diagnostics(rejected));
        Assert.Equal(new[] { "word/document.xml" }, rejected.RevisionFinalization.ChangedParts);
        Assert.True(rejected.RevisionFinalization.TrackingBefore);
        Assert.True(rejected.RevisionFinalization.TrackingAfter);
        using (var stream = new MemoryStream(rejected.File.ToByteArray()))
        using (var package = WordprocessingDocument.Open(stream, false))
        {
            var paragraphs = package.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().ToArray();
            Assert.Equal(string.Empty, paragraphs[0].InnerText);
            Assert.Equal("Removed wording", paragraphs[1].InnerText);
            Assert.IsType<W.Text>(paragraphs[1].Descendants<W.Run>().Single().ChildElements.Last());
            Assert.NotNull(paragraphs[1].Descendants<W.Bold>().SingleOrDefault());
            Assert.NotNull(package.MainDocumentPart.DocumentSettingsPart!.Settings!.GetFirstChild<W.TrackRevisions>());
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var importedAccepted = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportDocx,
            Family = ArtifactFamily.Document,
            File = accepted.File,
        });
        Assert.True(importedAccepted.Ok, Diagnostics(importedAccepted));
        Assert.False(importedAccepted.Artifact.Document.TrackRevisions);
        Assert.DoesNotContain(importedAccepted.Artifact.Document.Blocks, block => block.ContentCase == DocumentBlock.ContentOneofCase.Change);
        Assert.Equal(new[] { "Added wording", string.Empty }, importedAccepted.Artifact.Document.Blocks.Select(block => block.Paragraph.Text));
    }

    [Fact]
    public void RevisionFinalizationRejectsWrongHashNoChangesAndIrregularTopology()
    {
        var authored = Invoke(TrackedChangeExportRequest(includeDeletion: false));
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = authored.File.ToByteArray();
        var sourceHash = Convert.ToHexString(SHA256.HashData(source)).ToLowerInvariant();

        var wrongHash = Invoke(FinalizeRevisionRequest(source, new string('0', 64), DocumentRevisionFinalizationMode.Accept));
        Assert.False(wrongHash.Ok);
        Assert.Equal("document_source_hash_mismatch", Assert.Single(wrongHash.Diagnostics).Code);

        var noChanges = Invoke(ExportRequest(includeSecondParagraph: true));
        Assert.True(noChanges.Ok, Diagnostics(noChanges));
        var noChangeBytes = noChanges.File.ToByteArray();
        var noChangeHash = Convert.ToHexString(SHA256.HashData(noChangeBytes)).ToLowerInvariant();
        var noChangeResult = Invoke(FinalizeRevisionRequest(noChangeBytes, noChangeHash, DocumentRevisionFinalizationMode.Accept));
        Assert.False(noChangeResult.Ok);
        Assert.Equal("document_revisions_not_found", Assert.Single(noChangeResult.Diagnostics).Code);

        var irregular = AddSecondTrackedChangeRun(source);
        var irregularHash = Convert.ToHexString(SHA256.HashData(irregular)).ToLowerInvariant();
        var irregularResult = Invoke(FinalizeRevisionRequest(irregular, irregularHash, DocumentRevisionFinalizationMode.Accept));
        Assert.False(irregularResult.Ok);
        Assert.Equal("unsupported_document_revision_topology", Assert.Single(irregularResult.Diagnostics).Code);
    }

    [Fact]
    public void RevisionFinalizationFailsClosedWhenAnotherDocumentStoryContainsRevisions()
    {
        var authored = Invoke(TrackedChangeExportRequest(includeDeletion: false));
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = AddHeaderTrackedChange(authored.File.ToByteArray());
        var sourceHash = Convert.ToHexString(SHA256.HashData(source)).ToLowerInvariant();

        var response = Invoke(FinalizeRevisionRequest(source, sourceHash, DocumentRevisionFinalizationMode.Accept));
        Assert.False(response.Ok);
        var diagnostic = Assert.Single(response.Diagnostics);
        Assert.Equal("unsupported_document_revision_scope", diagnostic.Code);
        Assert.StartsWith("word/header", diagnostic.SourcePath, StringComparison.Ordinal);
    }

    private static CodecResponse Invoke(CodecRequest request) =>
        CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));

    private static string Diagnostics(CodecResponse response) =>
        string.Join("\n", response.Diagnostics.Select(item => $"{item.Code}: {item.Message}"));

    private static CodecRequest TrackedChangeExportRequest(bool includeDeletion = true, bool trackRevisions = false)
    {
        var document = new DocumentArtifact { Id = "document/tracked", Name = "Tracked changes", TrackRevisions = trackRevisions };
        document.Blocks.Add(new DocumentBlock
        {
            Id = "document/change/insert",
            StyleId = "Normal",
            Change = new DocumentChange
            {
                Type = DocumentChangeType.Insert,
                Text = "Added wording",
                Author = "Reviewer",
                Date = "2026-07-17T08:00:00Z",
            },
        });
        if (includeDeletion)
        {
            document.Blocks.Add(new DocumentBlock
            {
                Id = "document/change/delete",
                StyleId = "Normal",
                Change = new DocumentChange
                {
                    Type = DocumentChangeType.Delete,
                    Text = "Removed wording",
                    Author = "Reviewer",
                    Date = "2026-07-17T08:05:00Z",
                },
            });
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

    private static CodecRequest FinalizeRevisionRequest(
        byte[] source,
        string sourceHash,
        DocumentRevisionFinalizationMode mode,
        bool keepTracking = false) => new()
    {
        ProtocolVersion = CodecProtocol.ProtocolVersion,
        Operation = CodecOperation.FinalizeDocxRevisions,
        Family = ArtifactFamily.Document,
        File = ByteString.CopyFrom(source),
        RevisionFinalization = new DocumentRevisionFinalizationRequest
        {
            Mode = mode,
            KeepTracking = keepTracking,
            ExpectedSourceSha256 = sourceHash,
        },
    };

    private static byte[] AddTrackedChangeFormatting(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var run = document.MainDocumentPart!.Document!.Descendants<W.InsertedRun>().Single().Elements<W.Run>().Single();
            run.RunProperties = new W.RunProperties(new W.Bold());
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddAllTrackedChangeFormatting(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            foreach (var run in document.MainDocumentPart!.Document!.Descendants<W.InsertedRun>()
                         .SelectMany(change => change.Elements<W.Run>())
                         .Concat(document.MainDocumentPart.Document.Descendants<W.DeletedRun>()
                             .SelectMany(change => change.Elements<W.Run>())))
                run.RunProperties = new W.RunProperties(new W.Bold());
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddSecondTrackedChangeRun(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var insertion = document.MainDocumentPart!.Document!.Descendants<W.InsertedRun>().Single();
            insertion.Append(new W.Run(new W.Text(" continuation") { Space = SpaceProcessingModeValues.Preserve }));
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddHeaderTrackedChange(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var mainPart = document.MainDocumentPart!;
            var header = mainPart.AddNewPart<HeaderPart>();
            header.Header = new W.Header(new W.Paragraph(new W.InsertedRun(new W.Run(new W.Text("Header change")))
            {
                Id = "91",
                Author = "Reviewer",
            }));
            header.Header.Save();
            mainPart.Document!.Body!.Elements<W.SectionProperties>().Single().Append(new W.HeaderReference
            {
                Id = mainPart.GetIdOfPart(header),
                Type = W.HeaderFooterValues.Default,
            });
            mainPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static CodecRequest BibliographyExportRequest()
    {
        var document = new DocumentArtifact
        {
            Id = "document/bibliography",
            Name = "Bibliography fixture",
            Bibliography = new DocumentBibliography
            {
                SelectedStyle = "\\APASixthEditionOfficeOnline.xsl",
                StyleName = "APA",
                Uri = "https://example.test/styles/apa",
            },
        };
        var source = new DocumentBibliographySource
        {
            Id = "bibliography/AgentSource",
            Tag = "AgentSource",
            SourceType = "Book",
        };
        source.Authors.Add(new DocumentBibliographyPerson { First = "Ada", Last = "Lovelace" });
        source.Fields["title"] = "Sketch of the Analytical Engine";
        source.Fields["year"] = "1843";
        source.Fields["publisher"] = "Scientific Memoirs";
        source.Fields["guid"] = "{4D325651-1414-4D44-8BE3-4D44450E6C91}";
        document.Bibliography.Sources.Add(source);
        var citation = new DocumentBlock
        {
            Id = "document/citation",
            StyleId = "Normal",
            Citation = new DocumentCitation { Tag = source.Tag, Display = "(Lovelace, 1843)" },
        };
        document.Blocks.Add(citation);
        document.Bookmarks.Add(new DocumentBookmark
        {
            Id = "document/citation/bookmark",
            Name = "OpenOfficeCitation_AgentSource",
            TargetBlockId = citation.Id,
            EndTargetBlockId = citation.Id,
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

    private static byte[] MakeBibliographyContributorRoleIrregular(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var part = Assert.Single(document.MainDocumentPart!.CustomXmlParts);
            XNamespace bibliography = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography";
            XDocument xml;
            using (var input = part.GetStream(FileMode.Open, FileAccess.Read)) xml = XDocument.Load(input);
            var role = xml.Descendants(bibliography + "Author").Skip(1).First();
            role.Name = bibliography + "Editor";
            using var output = part.GetStream(FileMode.Create, FileAccess.Write);
            xml.Save(output, SaveOptions.DisableFormatting);
        }
        return stream.ToArray();
    }

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

    private static CodecRequest OfficeSkillProfileExportRequest()
    {
        var png = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=");
        var digest = Convert.ToHexString(SHA256.HashData(png)).ToLowerInvariant();
        var assetId = $"asset/document/image/{digest}";
        var document = new DocumentArtifact
        {
            Id = "document/office-skill-profile",
            Name = "Office skill profile",
            EvenAndOddHeaders = true,
            DefaultRunStyle = new DocumentRunFormatting
            {
                FontFamily = "Aptos",
                FontSizeHalfPoints = 22,
                ColorRgb = "202020",
            },
        };
        document.Styles.Add(new DocumentStyle
        {
            Id = "BriefLead",
            Name = "Brief lead",
            Type = DocumentStyleType.Paragraph,
            BasedOn = "Normal",
            RunFormat = new DocumentRunFormatting
            {
                FontFamily = "Aptos Display",
                FontSizeHalfPoints = 30,
                Bold = true,
            },
            ParagraphFormat = new DocumentParagraphFormatting
            {
                Alignment = "center",
                SpaceAfterTwips = 240,
                KeepNext = true,
            },
        });
        var lead = new DocumentBlock
        {
            Id = "document/lead",
            StyleId = "BriefLead",
            Paragraph = new DocumentParagraph
            {
                Text = "Styled office brief",
                Formatting = new DocumentParagraphFormatting
                {
                    Alignment = "center",
                    LeftIndentTwips = 360,
                    SpaceBeforeTwips = 120,
                    SpaceAfterTwips = 240,
                    LineSpacingTwips = 300,
                    LineSpacingRule = "auto",
                    KeepNext = true,
                },
            },
        };
        lead.Paragraph.Runs.Add(new DocumentRun
        {
            Text = lead.Paragraph.Text,
            Formatting = new DocumentRunFormatting
            {
                FontFamily = "Aptos Display",
                FontSizeHalfPoints = 30,
                ColorRgb = "315A83",
                CharacterSpacingTwips = 10,
                Bold = true,
                Italic = false,
                Underline = true,
            },
        });
        document.Blocks.Add(lead);
        document.Blocks.Add(new DocumentBlock
        {
            Id = "document/page-field",
            StyleId = "Normal",
            Field = new DocumentField { Instruction = "PAGE", Display = "1" },
        });
        document.Blocks.Add(new DocumentBlock
        {
            Id = "document/image",
            Name = "Chart preview",
            StyleId = "Normal",
            Image = new DocumentImage
            {
                AssetId = assetId,
                AltText = "Chart preview",
                WidthEmu = 1_905_000,
                HeightEmu = 952_500,
            },
        });
        document.Blocks.Add(new DocumentBlock
        {
            Id = "document/section",
            Section = new DocumentSection
            {
                BreakType = DocumentSectionBreak.Continuous,
                PageWidthTwips = 15_840,
                PageHeightTwips = 12_240,
                Landscape = true,
                MarginTopTwips = 720,
                MarginRightTwips = 900,
                MarginBottomTwips = 720,
                MarginLeftTwips = 900,
            },
        });
        var secondSection = new DocumentBlock
        {
            Id = "document/second-section-body",
            StyleId = "Normal",
            Paragraph = new DocumentParagraph { Text = "Second section" },
        };
        secondSection.Paragraph.Runs.Add(new DocumentRun { Text = secondSection.Paragraph.Text });
        document.Blocks.Add(secondSection);
        document.SectionSettings.Add(new DocumentSectionSettings
        {
            SectionIndex = 0,
            DifferentFirstPage = true,
        });
        document.Headers.Add(new DocumentHeaderFooter
        {
            Id = "document/header/default",
            Name = "Default header",
            StyleId = "Normal",
            Text = "OpenChestnut",
            Reference = DocumentHeaderFooterReference.Default,
            SectionIndex = 0,
            VariantActive = true,
        });
        document.Headers.Add(new DocumentHeaderFooter
        {
            Id = "document/header/first",
            Name = "First header",
            StyleId = "Normal",
            Text = "Office skill profile",
            Reference = DocumentHeaderFooterReference.First,
            SectionIndex = 0,
            VariantActive = true,
        });
        document.Footers.Add(new DocumentHeaderFooter
        {
            Id = "document/footer/even",
            Name = "Even page footer",
            StyleId = "Normal",
            Text = "1",
            FieldInstruction = "PAGE",
            Reference = DocumentHeaderFooterReference.Even,
            SectionIndex = 0,
            VariantActive = true,
        });
        var envelope = new ArtifactEnvelope
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Family = ArtifactFamily.Document,
            Document = document,
        };
        envelope.Assets.Add(new Asset
        {
            Id = assetId,
            FileName = "chart-preview.png",
            ContentType = "image/png",
            Data = ByteString.CopyFrom(png),
            Sha256 = digest,
        });
        return new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportDocx,
            Family = ArtifactFamily.Document,
            Artifact = envelope,
        };
    }

    private static CodecRequest ClassicCommentExportRequest()
    {
        var paragraph = new DocumentBlock
        {
            Id = "document/paragraph",
            StyleId = "Normal",
            Paragraph = new DocumentParagraph { Text = "Commented source paragraph" },
        };
        paragraph.Paragraph.Runs.Add(new DocumentRun { Text = paragraph.Paragraph.Text });
        var document = new DocumentArtifact { Id = "document/comments", Name = "Classic comments fixture" };
        document.Blocks.Add(paragraph);
        document.Comments.Add(new DocumentComment
        {
            Id = "comment/review",
            TargetBlockId = paragraph.Id,
            Author = "Reviewer",
            Initials = "RV",
            CreatedAt = "2026-07-16T08:00:00Z",
            Text = "Review this paragraph.",
        });
        document.Comments.Add(new DocumentComment
        {
            Id = "comment/link",
            TargetBlockId = paragraph.Id,
            Author = "Second reviewer",
            Initials = "SR",
            Text = "Keep the evidence link.",
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

    private static CodecRequest ModernCommentExportRequest()
    {
        var paragraph = new DocumentBlock
        {
            Id = "document/paragraph",
            StyleId = "Normal",
            Paragraph = new DocumentParagraph { Text = "Modern comment thread target" },
        };
        paragraph.Paragraph.Runs.Add(new DocumentRun { Text = paragraph.Paragraph.Text });
        var document = new DocumentArtifact { Id = "document/modern-comments", Name = "Modern comments fixture" };
        document.Blocks.Add(paragraph);
        document.Comments.Add(new DocumentComment
        {
            Id = "comment/root",
            TargetBlockId = paragraph.Id,
            Author = "Lead reviewer",
            Initials = "LR",
            CreatedAt = "2026-07-19T08:00:00Z",
            Text = "Please confirm the evidence.",
            Resolved = false,
            ParagraphId = "11111111",
            DurableId = "33333333",
            DateUtc = "2026-07-19T08:00:00Z",
            Person = new DocumentCommentPerson { ProviderId = "provider-a", UserId = "lead@example.test" },
            IntelligentPlaceholder = false,
        });
        document.Comments.Add(new DocumentComment
        {
            Id = "comment/reply",
            TargetBlockId = paragraph.Id,
            ParentCommentId = "comment/root",
            Author = "Second reviewer",
            Initials = "SR",
            CreatedAt = "2026-07-19T08:05:00Z",
            Text = "Evidence confirmed.",
            Resolved = false,
            ParagraphId = "22222222",
            DurableId = "44444444",
            DateUtc = "2026-07-19T08:05:00Z",
            Person = new DocumentCommentPerson { ProviderId = "provider-b", UserId = "second@example.test" },
            IntelligentPlaceholder = false,
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

    private static CodecRequest BookmarkExportRequest()
    {
        var document = new DocumentArtifact { Id = "document/bookmarks", Name = "Bookmark fixture" };
        document.Blocks.Add(new DocumentBlock
        {
            Id = "document/jump",
            StyleId = "Normal",
            Hyperlink = new DocumentHyperlink
            {
                Text = "Jump to target",
                InternalAnchor = "TargetBookmark",
            },
        });
        var target = new DocumentBlock
        {
            Id = "document/target",
            StyleId = "Normal",
            Paragraph = new DocumentParagraph { Text = "Target paragraph" },
        };
        target.Paragraph.Runs.Add(new DocumentRun { Text = target.Paragraph.Text });
        document.Blocks.Add(target);
        document.Bookmarks.Add(new DocumentBookmark
        {
            Id = "document/bookmark/target",
            Name = "TargetBookmark",
            TargetBlockId = target.Id,
            EndTargetBlockId = target.Id,
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

    private static CodecRequest NoteExportRequest()
    {
        var document = new DocumentArtifact { Id = "document/notes", Name = "Note fixture" };
        foreach (var (id, text) in new[]
                 {
                     ("document/note-target/footnote", "Paragraph with a footnote."),
                     ("document/note-target/endnote", "Paragraph with an endnote."),
                 })
        {
            var block = new DocumentBlock
            {
                Id = id,
                StyleId = "Normal",
                Paragraph = new DocumentParagraph { Text = text },
            };
            block.Paragraph.Runs.Add(new DocumentRun { Text = text });
            document.Blocks.Add(block);
        }
        document.Notes.Add(new DocumentNote
        {
            Id = "document/note/footnote",
            Kind = DocumentNoteKind.Footnote,
            TargetBlockId = document.Blocks[0].Id,
            Text = "Source-free footnote",
        });
        document.Notes.Add(new DocumentNote
        {
            Id = "document/note/endnote",
            Kind = DocumentNoteKind.Endnote,
            TargetBlockId = document.Blocks[1].Id,
            Text = "Source-free endnote",
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

    private static byte[] AddSecondFootnoteParagraph(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var footnote = document.MainDocumentPart!.FootnotesPart!.Footnotes!
                .Elements<W.Footnote>().Single(note => note.Id?.Value == 1);
            footnote.Append(new W.Paragraph(new W.Run(new W.Text("Second rich paragraph"))));
            document.MainDocumentPart.FootnotesPart.Footnotes.Save();
        }
        return stream.ToArray();
    }

    private static CodecRequest DirectNumberingExportRequest()
    {
        static DocumentBlock Item(
            string id,
            string text,
            uint numberingId,
            uint abstractNumberingId,
            uint level,
            string numberFormat,
            uint start,
            string levelText)
        {
            var block = new DocumentBlock
            {
                Id = id,
                StyleId = "Normal",
                Paragraph = new DocumentParagraph
                {
                    Text = text,
                    Numbering = new DocumentNumbering
                    {
                        NumberingId = numberingId,
                        AbstractNumberingId = abstractNumberingId,
                        Level = level,
                        NumberFormat = numberFormat,
                        Start = start,
                        LevelText = levelText,
                    },
                },
            };
            return block;
        }

        var document = new DocumentArtifact { Id = "document/direct-numbering", Name = "Direct numbering fixture" };
        document.Blocks.Add(Item("document/first", "First lettered item", 77, 9, 0, "upperLetter", 3, "%1)"));
        document.Blocks.Add(Item("document/nested", "Nested roman item", 77, 9, 2, "lowerRoman", 5, "%1.%2.%3."));
        document.Blocks.Add(Item("document/shared", "Second instance", 78, 9, 0, "upperLetter", 3, "%1)"));
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

    private static CodecRequest MergedTableExportRequest()
    {
        static DocumentTableCell Cell(
            uint gridColumn,
            uint columnSpan = 1,
            uint rowSpan = 1,
            DocumentTableVerticalMerge merge = DocumentTableVerticalMerge.Unspecified) => new()
        {
            GridColumn = gridColumn,
            ColumnSpan = columnSpan,
            RowSpan = rowSpan,
            VerticalMerge = merge,
            Editable = merge != DocumentTableVerticalMerge.Continue,
        };

        var table = new DocumentTable { GridColumns = 3 };
        table.Formatting = new DocumentTableFormatting
        {
            WidthDxa = 9000,
            IndentDxa = 240,
            CellMarginsDxa = new DocumentTableCellMargins { Top = 80, Bottom = 80, Start = 120, End = 120 },
            BorderColor = "445566",
            BorderSize = 8,
            HeaderFill = "E2E8F0",
        };
        table.Formatting.ColumnWidthsDxa.Add(3000);
        table.Formatting.ColumnWidthsDxa.Add(3000);
        table.Formatting.ColumnWidthsDxa.Add(3000);
        var first = new DocumentTableRow();
        first.Cells.Add("Merged owner");
        first.Cells.Add("Status");
        first.RichCells.Add(Cell(0, columnSpan: 2, rowSpan: 2, merge: DocumentTableVerticalMerge.Restart));
        first.RichCells.Add(Cell(2));
        table.Rows.Add(first);
        var second = new DocumentTableRow();
        second.Cells.Add(string.Empty);
        second.Cells.Add("Ready");
        second.RichCells.Add(Cell(0, columnSpan: 2, rowSpan: 0, merge: DocumentTableVerticalMerge.Continue));
        second.RichCells.Add(Cell(2));
        table.Rows.Add(second);
        var third = new DocumentTableRow();
        third.Cells.Add("Scope");
        third.Cells.Add("Complete");
        third.RichCells.Add(Cell(0));
        third.RichCells.Add(Cell(1, columnSpan: 2));
        table.Rows.Add(third);

        var document = new DocumentArtifact { Id = "document/merged-table", Name = "Merged table fixture" };
        document.Blocks.Add(new DocumentBlock { Id = "document/table", StyleId = "TableGrid", Table = table });
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

    private static CodecRequest TocExportRequest()
    {
        var document = new DocumentArtifact
        {
            Id = "document/toc",
            Name = "TOC fixture",
            UpdateFields = true,
        };
        document.Blocks.Add(new DocumentBlock
        {
            Id = "document/toc-field",
            StyleId = "Normal",
            Field = new DocumentField
            {
                Instruction = "TOC \\o \"1-3\" \\h \\z \\u",
                Display = "Refresh in Word",
                Complex = true,
            },
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

    private static byte[] AddSecondCommentBodyRun(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var comment = document.MainDocumentPart!.WordprocessingCommentsPart!.Comments!.Elements<W.Comment>().First();
            comment.Elements<W.Paragraph>().Single().Append(new W.Run(new W.Text(" preserved rich body")));
            document.MainDocumentPart.WordprocessingCommentsPart.Comments.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AliasSecondCommentId(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var mainPart = document.MainDocumentPart!;
            mainPart.WordprocessingCommentsPart!.Comments!.Elements<W.Comment>().ElementAt(1).Id = "00";
            foreach (var anchor in mainPart.Document!.Body!.Descendants<W.CommentRangeStart>().Where(item => item.Id?.Value == "1")) anchor.Id = "00";
            foreach (var anchor in mainPart.Document.Body.Descendants<W.CommentRangeEnd>().Where(item => item.Id?.Value == "1")) anchor.Id = "00";
            foreach (var reference in mainPart.Document.Body.Descendants<W.CommentReference>().Where(item => item.Id?.Value == "1")) reference.Id = "00";
            mainPart.WordprocessingCommentsPart.Comments.Save();
            mainPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] MakeModernReplyNested(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var commentsExPart = document.MainDocumentPart!.WordprocessingCommentsExPart!;
            var reply = commentsExPart.CommentsEx!.Elements<W15.CommentEx>().ElementAt(1);
            var paragraphId = reply.GetAttribute("paraId", "http://schemas.microsoft.com/office/word/2012/wordml").Value;
            reply.SetAttribute(new OpenXmlAttribute(
                "w15",
                "paraIdParent",
                "http://schemas.microsoft.com/office/word/2012/wordml",
                paragraphId));
            commentsExPart.CommentsEx.Save(commentsExPart);
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

    private static byte[] AddRecognizedTableFormattingResidual(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var rows = document.MainDocumentPart!.Document!.Body!.Elements<W.Table>().Single().Elements<W.TableRow>().ToArray();
            rows[1].PrependChild(new W.TableRowProperties(new W.TableRowHeight
            {
                Val = 480,
                HeightType = W.HeightRuleValues.AtLeast,
            }));
            rows[2].Elements<W.TableCell>().First().Elements<W.Paragraph>().Single()
                .PrependChild(new W.ParagraphProperties(new W.Justification { Val = W.JustificationValues.Center }));
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

    private static byte[] AddUnsupportedParagraphProperty(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var paragraph = document.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().ElementAt(1);
            var properties = paragraph.ParagraphProperties ?? paragraph.PrependChild(new W.ParagraphProperties());
            properties.Append(new W.WidowControl());
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddFragmentedReadOnlyParagraph(
        byte[] bytes,
        bool mixedFormatting = false,
        bool emptyRunGap = false)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var paragraph = document.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().ElementAt(1);
            var properties = paragraph.ParagraphProperties ?? paragraph.PrependChild(new W.ParagraphProperties());
            properties.Append(new W.WidowControl());
            var source = paragraph.Elements<W.Run>().Single();
            var first = FragmentRun(source, "Edit");
            var second = FragmentRun(source, "able paragraph");
            if (mixedFormatting)
            {
                var runProperties = second.RunProperties ?? second.PrependChild(new W.RunProperties());
                runProperties.Append(new W.Bold());
            }
            source.Remove();
            paragraph.Append(first);
            if (emptyRunGap) paragraph.Append(FragmentRun(first, string.Empty));
            paragraph.Append(second);
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddFragmentedTableCell(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var paragraph = document.MainDocumentPart!.Document!.Body!.Elements<W.Table>().Single()
                .Elements<W.TableRow>().Single().Elements<W.TableCell>().First().Elements<W.Paragraph>().Single();
            var source = paragraph.Elements<W.Run>().Single();
            source.InsertBeforeSelf(FragmentRun(source, "Rev"));
            source.InsertBeforeSelf(FragmentRun(source, "enue"));
            source.Remove();
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] WrapFirstTableCellRunInContentControl(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var paragraph = document.MainDocumentPart!.Document!.Body!.Elements<W.Table>().Single()
                .Elements<W.TableRow>().Single().Elements<W.TableCell>().First().Elements<W.Paragraph>().Single();
            var run = paragraph.Elements<W.Run>().Single();
            run.Remove();
            paragraph.Append(new W.SdtRun(
                new W.SdtProperties(new W.SdtId { Val = 8123 }, new W.Tag { Val = "protected-cell" }),
                new W.SdtContentRun(run)));
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static W.Run FragmentRun(W.Run source, string value)
    {
        var run = (W.Run)source.CloneNode(true);
        var text = run.Elements<W.Text>().Single();
        text.Text = value;
        text.Space = value.Length != value.Trim().Length ? SpaceProcessingModeValues.Preserve : null;
        return run;
    }

    private static string HashText(string value) =>
        Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(value))).ToLowerInvariant();

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

    private static byte[] AddDirectNumbering(byte[] bytes, bool addSecondRun = false, bool numberAllParagraphs = false, int level = 0)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var mainPart = document.MainDocumentPart!;
            var paragraphs = mainPart.Document!.Body!.Elements<W.Paragraph>().ToArray();
            var targets = numberAllParagraphs ? paragraphs : [paragraphs[1]];
            foreach (var paragraph in targets)
            {
                paragraph.ParagraphProperties = new W.ParagraphProperties(
                    new W.NumberingProperties(
                        new W.NumberingLevelReference { Val = level },
                        new W.NumberingId { Val = 77 }));
                var run = paragraph.Elements<W.Run>().Single();
                if (run.RunProperties is null) run.PrependChild(new W.RunProperties(new W.Italic()));
                else run.RunProperties.Append(new W.Italic());
            }
            if (addSecondRun) paragraphs[1].Append(new W.Run(new W.Text(" detail")));

            var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new W.Numbering(
                new W.AbstractNum(
                    new W.Level(
                        new W.StartNumberingValue { Val = 3 },
                        new W.NumberingFormat { Val = W.NumberFormatValues.UpperLetter },
                        new W.LevelText { Val = "%1)" }) { LevelIndex = level }) { AbstractNumberId = 9 },
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

            var stylesPart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
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

    private static byte[] AddUnusedStyleNumberingReference(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var mainPart = document.MainDocumentPart!;
            var stylesPart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles ??= new W.Styles();
            stylesPart.Styles.Append(new W.Style(
                new W.StyleName { Val = "Unused shared numbering reference" },
                new W.StyleParagraphProperties(
                    new W.NumberingProperties(new W.NumberingId { Val = 77 })))
            { Type = W.StyleValues.Paragraph, StyleId = "UnusedSharedList" });
            stylesPart.Styles.Save();
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

            var stylesPart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
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

    private static byte[] ExpandTocFieldAcrossParagraphs(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var body = document.MainDocumentPart!.Document!.Body!;
            var paragraph = body.Elements<W.Paragraph>().Single();
            var run = paragraph.Elements<W.Run>().Single();
            var begin = (W.FieldChar)run.ChildElements[0].CloneNode(true);
            var code = (W.FieldCode)run.ChildElements[1].CloneNode(true);
            var separate = (W.FieldChar)run.ChildElements[2].CloneNode(true);
            var end = (W.FieldChar)run.ChildElements[4].CloneNode(true);
            paragraph.InsertBeforeSelf(new W.Paragraph(new W.Run(begin, code, separate)));
            paragraph.InsertBeforeSelf(new W.Paragraph(new W.Run(new W.Text("First section\t1"))));
            paragraph.InsertBeforeSelf(new W.Paragraph(new W.Run(end)));
            paragraph.Remove();
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

    private static byte[] AddSecondParagraphToBlockContentControl(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var package = WordprocessingDocument.Open(stream, true))
        {
            var control = Assert.Single(package.MainDocumentPart!.Document!.Descendants<W.SdtBlock>());
            control.SdtContentBlock!.Append(new W.Paragraph(new W.Run(new W.Text("Unsupported second paragraph"))));
            package.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddDatePickerContentControl(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var properties = document.MainDocumentPart!.Document!.Descendants<W.SdtRun>().Single().SdtProperties!;
            properties.Append(new W.SdtContentDate());
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddCustomCheckboxSymbol(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var checkbox = document.MainDocumentPart!.Document!.Descendants<W14.SdtContentCheckBox>().Single();
            checkbox.GetFirstChild<W14.CheckedState>()!.Val = "F0FE";
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] ChangeDatePickerFormat(byte[] bytes, string format)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            document.MainDocumentPart!.Document!.Descendants<W.SdtContentDate>().Single().DateFormat!.Val = format;
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] DuplicateDropdownDisplayText(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var items = document.MainDocumentPart!.Document!.Descendants<W.SdtContentDropDownList>().Single().Elements<W.ListItem>().ToArray();
            items[1].DisplayText = items[0].DisplayText!.Value;
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

}
