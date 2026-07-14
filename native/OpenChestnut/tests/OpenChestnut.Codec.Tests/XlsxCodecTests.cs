using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace OpenChestnut.Codec.Tests;

public sealed class XlsxCodecTests
{
    [Fact]
    public void ProtocolRoundTripsMinimalWorkbook()
    {
        var request = ExportRequest();
        var export = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.True(export.Ok, string.Join("\n", export.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal("PK", System.Text.Encoding.ASCII.GetString(export.File.Span[..2]));

        using (var stream = new MemoryStream(export.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            var errors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document).ToArray();
            Assert.Empty(errors);
            var styles = document.WorkbookPart?.WorkbookStylesPart?.Stylesheet;
            Assert.NotNull(styles);
            Assert.Contains(styles!.NumberingFormats!.Elements<NumberingFormat>(), item => item.FormatCode?.Value == "0.000 \"units\"");
            var cells = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().ToDictionary(item => item.CellReference!.Value!);
            Assert.NotEqual(0U, cells["B1"].StyleIndex?.Value ?? 0U);
            Assert.NotEqual(0U, cells["B2"].StyleIndex?.Value ?? 0U);
        }

        var importRequest = new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportXlsx,
            Family = ArtifactFamily.Workbook,
            File = export.File,
        };
        var imported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(importRequest.ToByteArray()));
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal(WorkbookDateSystem._1904, imported.Artifact.Workbook.DateSystem);
        Assert.Equal("Summary", imported.Artifact.Workbook.Worksheets[0].Name);
        Assert.Collection(imported.Artifact.Workbook.Worksheets[0].Cells,
            cell => Assert.Equal("Quarter", cell.StringValue),
            cell =>
            {
                Assert.Equal(42.5, cell.NumberValue);
                Assert.Equal("0.000 \"units\"", cell.NumberFormatCode);
            },
            cell =>
            {
                Assert.Equal("=B1*2", cell.Formula);
                Assert.Equal("0.00%", cell.NumberFormatCode);
            });
    }

    [Fact]
    public void ProtocolReturnsStructuredBudgetFailure()
    {
        var export = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(ExportRequest().ToByteArray()));
        var request = new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportXlsx,
            Family = ArtifactFamily.Workbook,
            File = export.File,
            Limits = new CodecLimits { MaxInputBytes = 16 },
        };
        var response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("input_budget_exceeded", Assert.Single(response.Diagnostics).Code);
    }

    [Fact]
    public void ProtocolRejectsUnknownVersion()
    {
        var request = ExportRequest();
        request.ProtocolVersion = 99;
        var response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("unsupported_protocol_version", Assert.Single(response.Diagnostics).Code);
    }

    [Theory]
    [InlineData("../escape.xml", "unsafe_part_path")]
    [InlineData("xl/workbook.xml", "duplicate_part_path")]
    public void ImportRejectsUnsafeOrDuplicatePartPaths(string path, string expectedCode)
    {
        var firstExport = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(ExportRequest().ToByteArray()));
        var response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportXlsx,
            Family = ArtifactFamily.Workbook,
            File = ByteString.CopyFrom(AddEntry(firstExport.File.ToByteArray(), path)),
        }.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal(expectedCode, Assert.Single(response.Diagnostics).Code);
    }

    [Fact]
    public void ImportPreservesUnknownRelationshipFromHashBoundSourcePackage()
    {
        var firstExport = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(ExportRequest().ToByteArray()));
        var bytes = AddExternalWorkbookRelationship(firstExport.File.ToByteArray());
        var imported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportXlsx,
            Family = ArtifactFamily.Workbook,
            File = ByteString.CopyFrom(bytes),
        }.ToByteArray()));

        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var relationship = Assert.Single(imported.Artifact.OpaqueOpc.PackageRelationships);
        Assert.Equal("xl/workbook.xml", relationship.SourcePath);
        Assert.Equal("External", relationship.TargetMode);
        Assert.Equal("https://example.invalid/data", relationship.Target);
        Assert.False(imported.Artifact.OpaqueOpc.SourcePackage.Data.IsEmpty);
        Assert.All(imported.Artifact.OpaqueOpc.Parts, part => Assert.True(part.Data.IsEmpty));

        imported.Artifact.Workbook.Worksheets[0].Cells[1].NumberValue = 99;
        var preserved = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.True(preserved.Ok, string.Join("\n", preserved.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal("opaque_content_preserved", Assert.Single(preserved.Diagnostics).Code);
        Assert.Equal(ReadEntry(bytes, "xl/styles.xml"), ReadEntry(preserved.File.ToByteArray(), "xl/styles.xml"));

        var reimported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportXlsx,
            Family = ArtifactFamily.Workbook,
            File = preserved.File,
        }.ToByteArray()));
        Assert.True(reimported.Ok, string.Join("\n", reimported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal(99, reimported.Artifact.Workbook.Worksheets[0].Cells[1].NumberValue);
        Assert.Contains(reimported.Artifact.OpaqueOpc.PackageRelationships, item => item.Id == "rIdExternal" && item.Target == "https://example.invalid/data");

        imported.Artifact.OpaqueOpc.SourcePackage.Sha256 = new string('0', 64);
        var tampered = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.False(tampered.Ok);
        Assert.Equal("source_package_hash_mismatch", Assert.Single(tampered.Diagnostics).Code);

        imported.Artifact.OpaqueOpc.SourcePackage = new SourcePackageSnapshot();
        var rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("opaque_content_requires_preservation", Assert.Single(rejected.Diagnostics).Code);

        var lossy = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
            AllowLossy = true,
        }.ToByteArray()));
        Assert.True(lossy.Ok);
        Assert.Equal("opaque_content_discarded", Assert.Single(lossy.Diagnostics).Code);
    }

    [Fact]
    public void SourcePreservingExportRejectsInvalidOwnedMarkup()
    {
        var firstExport = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(ExportRequest().ToByteArray()));
        var imported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportXlsx,
            Family = ArtifactFamily.Workbook,
            File = ByteString.CopyFrom(AddInvalidWorksheetMarkup(firstExport.File.ToByteArray())),
        }.ToByteArray()));
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));

        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.False(exported.Ok);
        Assert.Equal("openxml_validation_failed", Assert.Single(exported.Diagnostics).Code);
    }

    [Fact]
    public void SourcePreservingNumberFormatEditClonesCellFormatAndKeepsUnmodeledProperties()
    {
        var firstExport = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(ExportRequest().ToByteArray()));
        var bytes = AddUnmodeledCellFormatProperties(firstExport.File.ToByteArray(), out var originalStyleIndex);
        var imported = Import(bytes);
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        imported.Artifact.Workbook.Worksheets[0].Cells[1].NumberFormatCode = "$#,##0.00";

        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));

        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            var styles = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet;
            var formats = styles!.CellFormats!.Elements<CellFormat>().ToArray();
            Assert.Equal(HorizontalAlignmentValues.Center, formats[originalStyleIndex].Alignment?.Horizontal?.Value);
            Assert.False(formats[originalStyleIndex].Protection?.Locked?.Value ?? true);
            var cell = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Single(item => item.CellReference?.Value == "B1");
            Assert.NotEqual(originalStyleIndex, checked((int)(cell.StyleIndex?.Value ?? 0)));
            var derived = formats[checked((int)cell.StyleIndex!.Value)];
            Assert.Equal(HorizontalAlignmentValues.Center, derived.Alignment?.Horizontal?.Value);
            Assert.False(derived.Protection?.Locked?.Value ?? true);
            Assert.Equal(formats[originalStyleIndex].FontId?.Value, derived.FontId?.Value);
            Assert.Equal(formats[originalStyleIndex].FillId?.Value, derived.FillId?.Value);
            Assert.Equal(formats[originalStyleIndex].BorderId?.Value, derived.BorderId?.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }

        var reimported = Import(exported.File.ToByteArray());
        Assert.True(reimported.Ok, string.Join("\n", reimported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal("$#,##0.00", reimported.Artifact.Workbook.Worksheets[0].Cells[1].NumberFormatCode);
    }

    [Fact]
    public void ProtocolRejectsInvalidNumberFormatCode()
    {
        var request = ExportRequest();
        request.Artifact.Workbook.Worksheets[0].Cells[1].NumberFormatCode = new string('x', XlsxNumberFormatCodec.MaxFormatCodeLength + 1);
        var response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_cell_number_format", Assert.Single(response.Diagnostics).Code);
    }

    [Fact]
    public void ImportRejectsMissingCellFormatReference()
    {
        var firstExport = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(ExportRequest().ToByteArray()));
        var response = Import(SetCellStyleIndex(firstExport.File.ToByteArray(), "B1", 999U));
        Assert.False(response.Ok);
        Assert.Equal("invalid_cell_number_format", Assert.Single(response.Diagnostics).Code);
    }

    [Fact]
    public void ImportRejectsMissingCustomNumberFormat()
    {
        var firstExport = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(ExportRequest().ToByteArray()));
        var response = Import(SetCellNumberFormatId(firstExport.File.ToByteArray(), "B1", 999U));
        Assert.False(response.Ok);
        Assert.Equal("invalid_cell_number_format", Assert.Single(response.Diagnostics).Code);
    }

    [Fact]
    public void ProtocolRoundTripsCompleteStaticCellStyleProfile()
    {
        var request = ExportRequest();
        var expected = FullStaticStyle();
        request.Artifact.Workbook.Worksheets[0].Cells[0].Style = expected;
        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));

        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            var styles = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            Assert.True((styles.Fonts?.Count?.Value ?? 0) >= 2);
            Assert.True((styles.Fills?.Count?.Value ?? 0) >= 3);
            Assert.True((styles.Borders?.Count?.Value ?? 0) >= 2);
            var cell = document.WorkbookPart.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Single(item => item.CellReference?.Value == "A1");
            var format = styles.CellFormats!.Elements<CellFormat>().ElementAt(checked((int)cell.StyleIndex!.Value));
            Assert.True(format.ApplyFont?.Value);
            Assert.True(format.ApplyFill?.Value);
            Assert.True(format.ApplyBorder?.Value);
            Assert.True(format.ApplyAlignment?.Value);
            Assert.True(format.ApplyProtection?.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }

        var imported = Import(exported.File.ToByteArray());
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal(expected, imported.Artifact.Workbook.Worksheets[0].Cells[0].Style);
    }

    [Fact]
    public void SourcePreservingStyleEditClonesResourcesAndRetainsResidualProperties()
    {
        var request = ExportRequest();
        request.Artifact.Workbook.Worksheets[0].Cells[0].Style = FullStaticStyle();
        var firstExport = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.True(firstExport.Ok);
        var bytes = AddUnmodeledStaticStyleProperties(firstExport.File.ToByteArray(), "A1", out var originalStyleIndex, out var originalFontIndex);
        var imported = Import(bytes);
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        imported.Artifact.Workbook.Worksheets[0].Cells[0].Style.Font.Bold = false;
        imported.Artifact.Workbook.Worksheets[0].Cells[0].Style.Fill.Foreground = new SpreadsheetColor { Rgb = "22C55E" };

        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));

        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            var styles = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            var formats = styles.CellFormats!.Elements<CellFormat>().ToArray();
            var fonts = styles.Fonts!.Elements<Font>().ToArray();
            Assert.True(formats[originalStyleIndex].QuotePrefix?.Value);
            Assert.Equal(FontSchemeValues.Minor, fonts[originalFontIndex].FontScheme?.Val?.Value);
            var cell = document.WorkbookPart.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Single(item => item.CellReference?.Value == "A1");
            var derived = formats[checked((int)cell.StyleIndex!.Value)];
            Assert.NotEqual(originalStyleIndex, checked((int)cell.StyleIndex.Value));
            Assert.True(derived.QuotePrefix?.Value);
            Assert.Equal(FontSchemeValues.Minor, fonts[checked((int)derived.FontId!.Value)].FontScheme?.Val?.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }

        var reimported = Import(exported.File.ToByteArray());
        var style = reimported.Artifact.Workbook.Worksheets[0].Cells[0].Style;
        Assert.False(style.Font.Bold);
        Assert.Equal("22C55E", style.Fill.Foreground.Rgb);
        Assert.Equal("0.000 \"units\"", reimported.Artifact.Workbook.Worksheets[0].Cells[1].NumberFormatCode);
    }

    [Fact]
    public void ProtocolRejectsInvalidStaticCellStyles()
    {
        var oversizedFont = ExportRequest();
        oversizedFont.Artifact.Workbook.Worksheets[0].Cells[0].Style = new CellStyleArtifact { Font = new SpreadsheetFontStyle { SizePoints = 500 } };
        var invalidFont = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(oversizedFont.ToByteArray()));
        Assert.False(invalidFont.Ok);
        Assert.Equal("invalid_cell_style", Assert.Single(invalidFont.Diagnostics).Code);

        var invalidColor = ExportRequest();
        invalidColor.Artifact.Workbook.Worksheets[0].Cells[0].Style = new CellStyleArtifact
        {
            Fill = new SpreadsheetFillStyle { PatternType = "solid", Foreground = new SpreadsheetColor { Theme = 12 } },
        };
        var colorResponse = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(invalidColor.ToByteArray()));
        Assert.False(colorResponse.Ok);
        Assert.Equal("invalid_cell_style", Assert.Single(colorResponse.Diagnostics).Code);

        var invalidBorder = ExportRequest();
        invalidBorder.Artifact.Workbook.Worksheets[0].Cells[0].Style = new CellStyleArtifact
        {
            Border = new SpreadsheetBorderStyle { Left = new SpreadsheetBorderEdgeStyle { Style = "triple" } },
        };
        var borderResponse = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(invalidBorder.ToByteArray()));
        Assert.False(borderResponse.Ok);
        Assert.Equal("invalid_cell_style", Assert.Single(borderResponse.Diagnostics).Code);
    }

    [Fact]
    public void SourceFreeExportDeduplicatesIdenticalStaticStyles()
    {
        var request = ExportRequest();
        request.Artifact.Workbook.Worksheets[0].Cells[0].Style = FullStaticStyle();
        request.Artifact.Workbook.Worksheets[0].Cells.Add(new CellArtifact { Row = 2, Column = 0, StringValue = "same", Style = FullStaticStyle() });
        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        using var stream = new MemoryStream(exported.File.ToByteArray());
        using var document = SpreadsheetDocument.Open(stream, false);
        var cells = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Where(item => item.CellReference?.Value is "A1" or "A3").ToArray();
        Assert.Equal(cells[0].StyleIndex?.Value, cells[1].StyleIndex?.Value);
    }

    [Fact]
    public void ProtocolRoundTripsCompleteWorkbookTheme()
    {
        var request = ExportRequest();
        var expected = CustomTheme();
        request.Artifact.Workbook.Theme = expected;
        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));

        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            Assert.NotNull(document.WorkbookPart!.ThemePart);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }

        var imported = Import(exported.File.ToByteArray());
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertTheme(expected, imported.Artifact.Workbook.Theme);
        Assert.True(imported.Artifact.Workbook.Theme.Source.Editable);
        Assert.Equal("xl/theme/theme1.xml", imported.Artifact.Workbook.Theme.Source.PartPath);
        Assert.Contains(imported.Artifact.OpaqueOpc.Parts, item => item.Path == "xl/theme/theme1.xml");
    }

    [Fact]
    public void SourcePreservingThemeEditRetainsSystemColorAndResidualContent()
    {
        var request = ExportRequest();
        request.Artifact.Workbook.Theme = CustomTheme();
        var first = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.True(first.Ok);
        var bytes = MutateTheme(first.File.ToByteArray(), unsupportedColor: false);
        var imported = Import(bytes);
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal("101010", imported.Artifact.Workbook.Theme.Dk1Rgb);
        Assert.True(imported.Artifact.Workbook.Theme.Source.Editable);
        imported.Artifact.Workbook.Theme.Name = "OpenChestnut Edited";
        imported.Artifact.Workbook.Theme.Accent2Rgb = "22C55E";

        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var xml = System.Text.Encoding.UTF8.GetString(ReadEntry(exported.File.ToByteArray(), "xl/theme/theme1.xml"));
        Assert.Contains("name=\"OpenChestnut Edited\"", xml);
        Assert.Contains("<a:sysClr val=\"windowText\" lastClr=\"101010\"", xml);
        Assert.Contains("probe value=\"preserve-me\"", xml);
        Assert.Contains("<a:fontScheme name=\"Office Clean Room\"", xml);
        Assert.Contains("<a:accent2><a:srgbClr val=\"22C55E\"", xml);

        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        var reimported = Import(exported.File.ToByteArray());
        Assert.Equal("OpenChestnut Edited", reimported.Artifact.Workbook.Theme.Name);
        Assert.Equal("101010", reimported.Artifact.Workbook.Theme.Dk1Rgb);
        Assert.Equal("22C55E", reimported.Artifact.Workbook.Theme.Accent2Rgb);
    }

    [Fact]
    public void UnsupportedSourceThemeIsPreservedAndReplacementFailsClosed()
    {
        var request = ExportRequest();
        request.Artifact.Workbook.Theme = CustomTheme();
        var first = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        var bytes = MutateTheme(first.File.ToByteArray(), unsupportedColor: true);
        var imported = Import(bytes);
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.False(imported.Artifact.Workbook.Theme.Source.Editable);
        Assert.Empty(imported.Artifact.Workbook.Theme.Accent1Rgb);

        var preserved = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.True(preserved.Ok, string.Join("\n", preserved.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal(ReadEntry(bytes, "xl/theme/theme1.xml"), ReadEntry(preserved.File.ToByteArray(), "xl/theme/theme1.xml"));

        var replacement = CustomTheme();
        replacement.Source = imported.Artifact.Workbook.Theme.Source.Clone();
        imported.Artifact.Workbook.Theme = replacement;
        var rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_workbook_theme", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void ProtocolRejectsInvalidThemeAndTamperedSourceBinding()
    {
        var incomplete = ExportRequest();
        incomplete.Artifact.Workbook.Theme = new SpreadsheetThemeArtifact { Name = "Incomplete", Accent1Rgb = "0F172A" };
        var missing = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(incomplete.ToByteArray()));
        Assert.False(missing.Ok);
        Assert.Equal("invalid_workbook_theme", Assert.Single(missing.Diagnostics).Code);

        var invalid = ExportRequest();
        invalid.Artifact.Workbook.Theme = CustomTheme();
        invalid.Artifact.Workbook.Theme.Accent1Rgb = "not-rgb";
        var malformed = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(invalid.ToByteArray()));
        Assert.False(malformed.Ok);
        Assert.Equal("invalid_workbook_theme", Assert.Single(malformed.Diagnostics).Code);

        var valid = ExportRequest();
        valid.Artifact.Workbook.Theme = CustomTheme();
        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(valid.ToByteArray()));
        var imported = Import(exported.File.ToByteArray());
        imported.Artifact.Workbook.Theme.Source.XmlSha256 = new string('0', 64);
        var tampered = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.False(tampered.Ok);
        Assert.Equal("invalid_workbook_theme", Assert.Single(tampered.Diagnostics).Code);
    }

    [Fact]
    public void ProtocolAuthorsAndImportsBoundedWorksheetTable()
    {
        var request = TableExportRequest();
        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            var worksheetPart = document.WorkbookPart!.WorksheetParts.Single();
            Assert.Single(worksheetPart.TableDefinitionParts);
            Assert.Single(worksheetPart.Worksheet!.GetFirstChild<TableParts>()!.Elements<TablePart>());
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }
        var imported = Import(exported.File.ToByteArray());
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var table = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Tables);
        Assert.Equal("SalesTable", table.Name);
        Assert.Equal("A1:B3", table.Reference);
        Assert.Equal(["Region", "Revenue"], table.ColumnNames);
        Assert.Equal("TableStyleMedium4", table.StyleName);
        Assert.True(table.Source.Editable);
        Assert.Equal("xl/worksheets/sheet1.xml", table.Source.WorksheetPartPath);
        Assert.Equal("xl/tables/table1.xml", table.Source.TablePartPath);
    }

    [Fact]
    public void SourcePreservingWorksheetTableEditKeepsPartIdentity()
    {
        var first = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(TableExportRequest().ToByteArray()));
        var imported = Import(first.File.ToByteArray());
        var table = imported.Artifact.Workbook.Worksheets[0].Tables[0];
        var path = table.Source.TablePartPath;
        var relationshipId = table.Source.RelationshipId;
        table.Name = "RevenueTable";
        table.StyleName = "TableStyleMedium9";
        table.ShowFirstColumn = true;
        table.ShowColumnStripes = true;
        table.ColumnNames[1] = "Net Revenue";
        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var reimported = Import(exported.File.ToByteArray());
        var edited = Assert.Single(reimported.Artifact.Workbook.Worksheets[0].Tables);
        Assert.Equal("RevenueTable", edited.Name);
        Assert.Equal("TableStyleMedium9", edited.StyleName);
        Assert.True(edited.ShowFirstColumn);
        Assert.True(edited.ShowColumnStripes);
        Assert.Equal("Net Revenue", edited.ColumnNames[1]);
        Assert.Equal(path, edited.Source.TablePartPath);
        Assert.Equal(relationshipId, edited.Source.RelationshipId);
    }

    [Fact]
    public void UnsupportedWorksheetTableIsPreservedAndReplacementFailsClosed()
    {
        var first = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(TableExportRequest().ToByteArray()));
        var bytes = MutateTable(first.File.ToByteArray());
        var imported = Import(bytes);
        var table = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Tables);
        Assert.False(table.Source.Editable);
        Assert.Empty(table.Name);
        var preserved = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.True(preserved.Ok, string.Join("\n", preserved.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal(ReadEntry(bytes, "xl/tables/table1.xml"), ReadEntry(preserved.File.ToByteArray(), "xl/tables/table1.xml"));
        var replacement = TableArtifact();
        replacement.Source = table.Source.Clone();
        imported.Artifact.Workbook.Worksheets[0].Tables[0] = replacement;
        var rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void ProtocolRejectsInvalidWorksheetTableProfiles()
    {
        var request = TableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].ColumnNames.RemoveAt(1);
        var response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);

        request = TableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].Name = "A1";
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);
    }

    [Fact]
    public void ProtocolRoundTripsSharedAndLegacyArrayFormulaTopology()
    {
        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(FormulaExportRequest().ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));

        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            var cells = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().ToDictionary(item => item.CellReference!.Value!);
            Assert.Equal(CellFormulaValues.Shared, cells["C1"].CellFormula!.FormulaType!.Value);
            Assert.Equal(7U, cells["C1"].CellFormula!.SharedIndex!.Value);
            Assert.Equal("C1:C2", cells["C1"].CellFormula!.Reference!.Value);
            Assert.Equal("LOG10(A1)+B1", cells["C1"].CellFormula!.Text);
            Assert.Equal(CellFormulaValues.Shared, cells["C2"].CellFormula!.FormulaType!.Value);
            Assert.Equal(7U, cells["C2"].CellFormula!.SharedIndex!.Value);
            Assert.Null(cells["C2"].CellFormula!.Reference);
            Assert.Empty(cells["C2"].CellFormula!.Text);
            Assert.Equal(CellFormulaValues.Array, cells["E1"].CellFormula!.FormulaType!.Value);
            Assert.Equal("E1:E2", cells["E1"].CellFormula!.Reference!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }

        var imported = Import(exported.File.ToByteArray());
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var cellsByAddress = imported.Artifact.Workbook.Worksheets[0].Cells.ToDictionary(cell => $"{cell.Row}:{cell.Column}");
        Assert.Equal("=LOG10(A1)+B1", cellsByAddress["0:2"].Formula);
        Assert.Equal(CellFormulaKind.Shared, cellsByAddress["0:2"].FormulaMetadata.Kind);
        Assert.Equal(7U, cellsByAddress["0:2"].FormulaMetadata.SharedIndex);
        Assert.Equal("C1:C2", cellsByAddress["0:2"].FormulaMetadata.Reference);
        Assert.Equal("=LOG10(A2)+B2", cellsByAddress["1:2"].Formula);
        Assert.Equal(CellFormulaKind.Shared, cellsByAddress["1:2"].FormulaMetadata.Kind);
        Assert.Equal("=A1:A2*B1:B2", cellsByAddress["0:4"].Formula);
        Assert.Equal(CellFormulaKind.Array, cellsByAddress["0:4"].FormulaMetadata.Kind);
        Assert.Equal("E1:E2", cellsByAddress["0:4"].FormulaMetadata.Reference);
    }

    [Fact]
    public void SourcePreservingCachedValueAndNumberFormatEditsKeepFormulaXmlExact()
    {
        var first = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(FormulaExportRequest().ToByteArray()));
        var before = ReadFormulaXml(first.File.ToByteArray());
        var imported = Import(first.File.ToByteArray());
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var cells = imported.Artifact.Workbook.Worksheets[0].Cells.ToDictionary(cell => $"{cell.Row}:{cell.Column}");
        cells["1:2"].NumberValue = 99;
        cells["0:2"].NumberFormatCode = "0.00";

        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var after = ReadFormulaXml(exported.File.ToByteArray());
        Assert.Equal(before["C1"], after["C1"]);
        Assert.Equal(before["C2"], after["C2"]);
        Assert.Equal(before["E1"], after["E1"]);
    }

    [Fact]
    public void SourcePreservingExportCanDetachACompleteSharedGroupToNormalFormulas()
    {
        var first = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(FormulaExportRequest().ToByteArray()));
        var imported = Import(first.File.ToByteArray());
        Assert.True(imported.Ok);
        foreach (var cell in imported.Artifact.Workbook.Worksheets[0].Cells.Where(cell => cell.FormulaMetadata?.Kind == CellFormulaKind.Shared))
            cell.FormulaMetadata = null;

        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        using var stream = new MemoryStream(exported.File.ToByteArray());
        using var document = SpreadsheetDocument.Open(stream, false);
        var cells = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().ToDictionary(item => item.CellReference!.Value!);
        Assert.Null(cells["C1"].CellFormula!.FormulaType);
        Assert.Equal("LOG10(A1)+B1", cells["C1"].CellFormula!.Text);
        Assert.Null(cells["C2"].CellFormula!.FormulaType);
        Assert.Equal("LOG10(A2)+B2", cells["C2"].CellFormula!.Text);
    }

    [Fact]
    public void ProtocolRejectsIncompleteSharedFormulaGroup()
    {
        var request = FormulaExportRequest();
        request.Artifact.Workbook.Worksheets[0].Cells.Remove(request.Artifact.Workbook.Worksheets[0].Cells.Single(cell => cell.Row == 1 && cell.Column == 2));
        var response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_cell_formula", Assert.Single(response.Diagnostics).Code);

        var nested = FormulaExportRequest();
        nested.Artifact.Workbook.Worksheets[0].Cells.Add(new CellArtifact { Row = 1, Column = 4, Formula = "=1+1", NumberValue = 2 });
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(nested.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_cell_formula", Assert.Single(response.Diagnostics).Code);

        var oversized = FormulaExportRequest();
        oversized.Artifact.Workbook.Worksheets[0].Cells.Single(cell => cell.FormulaMetadata?.Kind == CellFormulaKind.Array).FormulaMetadata.Reference = "A1:XFD1048576";
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(oversized.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_cell_formula", Assert.Single(response.Diagnostics).Code);
    }

    [Fact]
    public void ImportRejectsDataTableFormulaAndFormulaEditRejectsUnmodeledAttributes()
    {
        var first = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(FormulaExportRequest().ToByteArray()));
        var dataTable = Import(SetFormulaType(first.File.ToByteArray(), "C1", CellFormulaValues.DataTable));
        Assert.False(dataTable.Ok);
        Assert.Equal("unsupported_cell_formula", Assert.Single(dataTable.Diagnostics).Code);

        var nestedArrayFormula = Import(AddFormulaCell(first.File.ToByteArray(), "E2", "1+1"));
        Assert.False(nestedArrayFormula.Ok);
        Assert.Equal("invalid_cell_formula", Assert.Single(nestedArrayFormula.Diagnostics).Code);

        var attributed = Import(SetFormulaCalculateCell(first.File.ToByteArray(), "C1"));
        Assert.True(attributed.Ok, string.Join("\n", attributed.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var cells = attributed.Artifact.Workbook.Worksheets[0].Cells.ToDictionary(cell => $"{cell.Row}:{cell.Column}");
        cells["0:2"].Formula = "=A1-B1";
        cells["1:2"].Formula = "=A2-B2";
        var rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = attributed.Artifact,
        }.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_cell_formula_edit", Assert.Single(rejected.Diagnostics).Code);
    }

    private static CodecRequest ExportRequest()
    {
        var sheet = new WorksheetArtifact
        {
            Id = "worksheet/summary",
            Name = "Summary",
            ShowGridLines = false,
            FreezePane = new FreezePane { Rows = 1, Columns = 1, TopLeftCell = "B2" },
        };
        sheet.Cells.Add(new CellArtifact { Row = 0, Column = 0, StringValue = "Quarter" });
        sheet.Cells.Add(new CellArtifact { Row = 0, Column = 1, NumberValue = 42.5, NumberFormatCode = "0.000 \"units\"" });
        sheet.Cells.Add(new CellArtifact { Row = 1, Column = 1, Formula = "=B1*2", NumberValue = 85, NumberFormatCode = "0.00%" });
        sheet.ColumnDimensions.Add(new ColumnDimension { Column = 0, Width = 18, BestFit = true });
        sheet.RowDimensions.Add(new RowDimension { Row = 0, Height = 24 });
        sheet.MergedRanges.Add("A3:B3");
        var workbook = new WorkbookArtifact { Id = "workbook/test", DateSystem = WorkbookDateSystem._1904 };
        workbook.Worksheets.Add(sheet);
        return new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = new ArtifactEnvelope
            {
                ProtocolVersion = CodecProtocol.ProtocolVersion,
                Family = ArtifactFamily.Workbook,
                Workbook = workbook,
            },
        };
    }

    private static CodecRequest FormulaExportRequest()
    {
        var sheet = new WorksheetArtifact { Id = "worksheet/formulas", Name = "Formulas", ShowGridLines = true };
        sheet.Cells.Add(new CellArtifact { Row = 0, Column = 0, NumberValue = 2 });
        sheet.Cells.Add(new CellArtifact { Row = 0, Column = 1, NumberValue = 3 });
        sheet.Cells.Add(new CellArtifact
        {
            Row = 0, Column = 2, Formula = "=LOG10(A1)+B1", NumberValue = 5,
            FormulaMetadata = new CellFormulaMetadata { Kind = CellFormulaKind.Shared, SharedIndex = 7, Reference = "C1:C2" },
        });
        sheet.Cells.Add(new CellArtifact
        {
            Row = 0, Column = 4, Formula = "=A1:A2*B1:B2", NumberValue = 6,
            FormulaMetadata = new CellFormulaMetadata { Kind = CellFormulaKind.Array, Reference = "E1:E2" },
        });
        sheet.Cells.Add(new CellArtifact { Row = 1, Column = 0, NumberValue = 4 });
        sheet.Cells.Add(new CellArtifact { Row = 1, Column = 1, NumberValue = 5 });
        sheet.Cells.Add(new CellArtifact
        {
            Row = 1, Column = 2, Formula = "=LOG10(A2)+B2", NumberValue = 9,
            FormulaMetadata = new CellFormulaMetadata { Kind = CellFormulaKind.Shared, SharedIndex = 7, Reference = "C1:C2" },
        });
        var workbook = new WorkbookArtifact { Id = "workbook/formulas", DateSystem = WorkbookDateSystem._1900 };
        workbook.Worksheets.Add(sheet);
        return new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = new ArtifactEnvelope
            {
                ProtocolVersion = CodecProtocol.ProtocolVersion,
                Family = ArtifactFamily.Workbook,
                Workbook = workbook,
            },
        };
    }

    private static CodecRequest TableExportRequest()
    {
        var request = ExportRequest();
        var sheet = request.Artifact.Workbook.Worksheets[0];
        sheet.Cells.Clear();
        sheet.Cells.Add(new CellArtifact { Row = 0, Column = 0, StringValue = "Region" });
        sheet.Cells.Add(new CellArtifact { Row = 0, Column = 1, StringValue = "Revenue" });
        sheet.Cells.Add(new CellArtifact { Row = 1, Column = 0, StringValue = "North" });
        sheet.Cells.Add(new CellArtifact { Row = 1, Column = 1, NumberValue = 120 });
        sheet.Cells.Add(new CellArtifact { Row = 2, Column = 0, StringValue = "South" });
        sheet.Cells.Add(new CellArtifact { Row = 2, Column = 1, NumberValue = 80 });
        sheet.Tables.Add(TableArtifact());
        return request;
    }

    private static SpreadsheetTableArtifact TableArtifact()
    {
        var table = new SpreadsheetTableArtifact
        {
            Id = "table/sales",
            Name = "SalesTable",
            Reference = "A1:B3",
            HasHeaders = true,
            ShowTotals = false,
            ShowFilterButton = true,
            StyleName = "TableStyleMedium4",
            ShowRowStripes = true,
        };
        table.ColumnNames.Add("Region");
        table.ColumnNames.Add("Revenue");
        return table;
    }

    private static byte[] MutateTable(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("xl/tables/table1.xml") ?? throw new InvalidOperationException("Worksheet table is missing.");
            XDocument table;
            using (var reader = new StreamReader(entry.Open())) table = XDocument.Parse(reader.ReadToEnd(), LoadOptions.PreserveWhitespace);
            table.Root!.SetAttributeValue("published", "0");
            entry.Delete();
            var replacement = archive.CreateEntry("xl/tables/table1.xml");
            using var writer = new StreamWriter(replacement.Open());
            writer.Write(table.ToString(SaveOptions.DisableFormatting));
        }
        return stream.ToArray();
    }

    private static SpreadsheetThemeArtifact CustomTheme() => new()
    {
        Name = "OpenChestnut Theme",
        Dk1Rgb = "101820", Lt1Rgb = "F8FAFC", Dk2Rgb = "1E3A5F", Lt2Rgb = "E2E8F0",
        Accent1Rgb = "0F766E", Accent2Rgb = "C2410C", Accent3Rgb = "4D7C0F", Accent4Rgb = "7E22CE",
        Accent5Rgb = "0369A1", Accent6Rgb = "BE123C", HlinkRgb = "1D4ED8", FolHlinkRgb = "7E22CE",
    };

    private static void AssertTheme(SpreadsheetThemeArtifact expected, SpreadsheetThemeArtifact actual)
    {
        Assert.Equal(expected.Name, actual.Name);
        Assert.Equal(expected.Dk1Rgb, actual.Dk1Rgb); Assert.Equal(expected.Lt1Rgb, actual.Lt1Rgb);
        Assert.Equal(expected.Dk2Rgb, actual.Dk2Rgb); Assert.Equal(expected.Lt2Rgb, actual.Lt2Rgb);
        Assert.Equal(expected.Accent1Rgb, actual.Accent1Rgb); Assert.Equal(expected.Accent2Rgb, actual.Accent2Rgb);
        Assert.Equal(expected.Accent3Rgb, actual.Accent3Rgb); Assert.Equal(expected.Accent4Rgb, actual.Accent4Rgb);
        Assert.Equal(expected.Accent5Rgb, actual.Accent5Rgb); Assert.Equal(expected.Accent6Rgb, actual.Accent6Rgb);
        Assert.Equal(expected.HlinkRgb, actual.HlinkRgb); Assert.Equal(expected.FolHlinkRgb, actual.FolHlinkRgb);
    }

    private static byte[] MutateTheme(byte[] bytes, bool unsupportedColor)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("xl/theme/theme1.xml") ?? throw new InvalidOperationException("Workbook theme is missing.");
            XDocument theme;
            using (var reader = new StreamReader(entry.Open())) theme = XDocument.Parse(reader.ReadToEnd(), LoadOptions.PreserveWhitespace);
            XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace probe = "urn:openchestnut:test";
            var scheme = theme.Root!.Element(drawing + "themeElements")!.Element(drawing + "clrScheme")!;
            scheme.Element(drawing + "dk1")!.Elements().Single().ReplaceWith(new XElement(drawing + "sysClr", new XAttribute("val", "windowText"), new XAttribute("lastClr", "101010")));
            if (unsupportedColor)
                scheme.Element(drawing + "accent1")!.Elements().Single().ReplaceWith(new XElement(drawing + "hslClr", new XAttribute("hue", 0), new XAttribute("sat", 100000), new XAttribute("lum", 50000)));
            theme.Root.Add(new XElement(drawing + "extLst", new XElement(drawing + "ext", new XAttribute("uri", "urn:openchestnut:test"), new XElement(probe + "probe", new XAttribute("value", "preserve-me")))));
            entry.Delete();
            var replacement = archive.CreateEntry("xl/theme/theme1.xml");
            using var writer = new StreamWriter(replacement.Open());
            writer.Write(theme.ToString(SaveOptions.DisableFormatting));
        }
        return stream.ToArray();
    }

    private static CellStyleArtifact FullStaticStyle() => new()
    {
        Font = new SpreadsheetFontStyle
        {
            Bold = true, Italic = false, Underline = "double", Strike = true,
            Color = new SpreadsheetColor { Theme = 4, Tint = 0.25 }, SizePoints = 14, Name = "Aptos Display",
        },
        Fill = new SpreadsheetFillStyle
        {
            PatternType = "solid",
            Foreground = new SpreadsheetColor { Rgb = "0F172A" },
        },
        Border = new SpreadsheetBorderStyle
        {
            Left = new SpreadsheetBorderEdgeStyle { Style = "thin", Color = new SpreadsheetColor { Indexed = 8 } },
            Right = new SpreadsheetBorderEdgeStyle { Style = "thin", Color = new SpreadsheetColor { Indexed = 8 } },
            Top = new SpreadsheetBorderEdgeStyle { Style = "thin", Color = new SpreadsheetColor { Indexed = 8 } },
            Bottom = new SpreadsheetBorderEdgeStyle { Style = "double", Color = new SpreadsheetColor { Automatic = true } },
            Diagonal = new SpreadsheetBorderEdgeStyle { Style = "dashed", Color = new SpreadsheetColor { Rgb = "EF4444" } },
            Start = new SpreadsheetBorderEdgeStyle { Style = "hair", Color = new SpreadsheetColor { Theme = 5 } },
            End = new SpreadsheetBorderEdgeStyle { Style = "medium", Color = new SpreadsheetColor { Rgb = "22C55E" } },
            Horizontal = new SpreadsheetBorderEdgeStyle { Style = "dotted", Color = new SpreadsheetColor { Indexed = 9 } },
            Vertical = new SpreadsheetBorderEdgeStyle { Style = "dashDot", Color = new SpreadsheetColor { Automatic = true } },
            DiagonalUp = true,
            DiagonalDown = true,
            Outline = false,
        },
        Alignment = new SpreadsheetAlignmentStyle
        {
            Horizontal = "center", Vertical = "bottom", WrapText = true, TextRotation = 45,
            Indent = 2, ShrinkToFit = false, ReadingOrder = 1,
        },
        Protection = new SpreadsheetProtectionStyle { Locked = false, Hidden = true },
    };

    private static CodecResponse Import(byte[] bytes) => CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
    {
        ProtocolVersion = CodecProtocol.ProtocolVersion,
        Operation = CodecOperation.ImportXlsx,
        Family = ArtifactFamily.Workbook,
        File = ByteString.CopyFrom(bytes),
    }.ToByteArray()));

    private static byte[] AddUnmodeledCellFormatProperties(byte[] bytes, out int styleIndex)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var cell = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Single(item => item.CellReference?.Value == "B1");
            styleIndex = checked((int)(cell.StyleIndex?.Value ?? 0));
            var format = document.WorkbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>().ElementAt(styleIndex);
            format.Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center };
            format.ApplyAlignment = true;
            format.Protection = new Protection { Locked = false };
            format.ApplyProtection = true;
            document.WorkbookPart.WorkbookStylesPart.Stylesheet!.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddUnmodeledStaticStyleProperties(byte[] bytes, string reference, out int styleIndex, out int fontIndex)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var cell = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Single(item => item.CellReference?.Value == reference);
            styleIndex = checked((int)(cell.StyleIndex?.Value ?? 0));
            var styles = document.WorkbookPart.WorkbookStylesPart!.Stylesheet!;
            var format = styles.CellFormats!.Elements<CellFormat>().ElementAt(styleIndex);
            format.QuotePrefix = true;
            fontIndex = checked((int)(format.FontId?.Value ?? 0));
            var font = styles.Fonts!.Elements<Font>().ElementAt(fontIndex);
            font.FontScheme = new FontScheme { Val = FontSchemeValues.Minor };
            styles.Save();
        }
        return stream.ToArray();
    }

    private static byte[] SetCellStyleIndex(byte[] bytes, string reference, uint styleIndex)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            var cell = worksheet.Descendants<Cell>().Single(item => item.CellReference?.Value == reference);
            cell.StyleIndex = styleIndex;
            worksheet.Save();
        }
        return stream.ToArray();
    }

    private static byte[] SetCellNumberFormatId(byte[] bytes, string reference, uint numberFormatId)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var cell = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Single(item => item.CellReference?.Value == reference);
            var format = document.WorkbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>().ElementAt(checked((int)(cell.StyleIndex?.Value ?? 0)));
            format.NumberFormatId = numberFormatId;
            format.ApplyNumberFormat = true;
            document.WorkbookPart.WorkbookStylesPart.Stylesheet.Save();
        }
        return stream.ToArray();
    }

    private static Dictionary<string, string> ReadFormulaXml(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        return document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>()
            .Where(cell => cell.CellFormula is not null)
            .ToDictionary(cell => cell.CellReference!.Value!, cell => cell.CellFormula!.OuterXml);
    }

    private static byte[] SetFormulaType(byte[] bytes, string reference, CellFormulaValues type)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            var formula = worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == reference).CellFormula!;
            formula.FormulaType = type;
            worksheet.Save();
        }
        return stream.ToArray();
    }

    private static byte[] SetFormulaCalculateCell(byte[] bytes, string reference)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            var formula = worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == reference).CellFormula!;
            formula.CalculateCell = true;
            worksheet.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddFormulaCell(byte[] bytes, string reference, string formulaText)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            var rowIndex = uint.Parse(new string(reference.SkipWhile(char.IsLetter).ToArray()));
            var row = worksheet.Descendants<Row>().Single(item => item.RowIndex?.Value == rowIndex);
            row.Append(new Cell { CellReference = reference, CellFormula = new CellFormula(formulaText), CellValue = new CellValue("2") });
            worksheet.Save();
        }
        return stream.ToArray();
    }

    private static byte[] ReadEntry(byte[] bytes, string path)
    {
        using var stream = new MemoryStream(bytes);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        using var entry = (archive.GetEntry(path) ?? throw new InvalidOperationException($"Missing package entry {path}.")).Open();
        using var output = new MemoryStream();
        entry.CopyTo(output);
        return output.ToArray();
    }

    private static byte[] AddExternalWorkbookRelationship(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("xl/_rels/workbook.xml.rels") ?? throw new InvalidOperationException("Workbook relationships are missing.");
            string xml;
            using (var reader = new StreamReader(entry.Open())) xml = reader.ReadToEnd();
            entry.Delete();
            var replacement = archive.CreateEntry("xl/_rels/workbook.xml.rels");
            using var writer = new StreamWriter(replacement.Open());
            writer.Write(xml.Replace("</Relationships>", "<Relationship Id=\"rIdExternal\" Type=\"urn:open-office-artifact-tool:test\" Target=\"https://example.invalid/data\" TargetMode=\"External\"/></Relationships>", StringComparison.Ordinal));
        }
        return stream.ToArray();
    }

    private static byte[] AddEntry(byte[] bytes, string path)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        using (var writer = new StreamWriter(archive.CreateEntry(path).Open()))
            writer.Write("<probe/>");
        return stream.ToArray();
    }

    private static byte[] AddInvalidWorksheetMarkup(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("xl/worksheets/sheet1.xml") ?? throw new InvalidOperationException("Worksheet is missing.");
            string xml;
            using (var reader = new StreamReader(entry.Open())) xml = reader.ReadToEnd();
            var closing = xml.LastIndexOf("</", StringComparison.Ordinal);
            entry.Delete();
            var replacement = archive.CreateEntry("xl/worksheets/sheet1.xml");
            using var writer = new StreamWriter(replacement.Open());
            writer.Write(xml.Insert(closing, "<x:notARealWorksheetChild/>"));
        }
        return stream.ToArray();
    }
}
