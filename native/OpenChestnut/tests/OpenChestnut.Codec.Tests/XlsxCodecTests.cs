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
    public void ImportsAndEditsSourceBoundWorksheetQueryTableGraph()
    {
        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(TableExportRequest().ToByteArray()));
        var source = AddQueryTableGraph(authored.File.ToByteArray());
        AssertOffice2021Valid(source);
        var connectionXml = ReadEntry(source, "xl/connections.xml");
        var relationshipXml = ReadEntry(source, "xl/tables/_rels/table1.xml.rels");

        var imported = Import(source);
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var table = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Tables);
        Assert.True(table.Source.Editable);
        var query = Assert.IsType<SpreadsheetTableQueryArtifact>(table.QueryTable);
        Assert.Equal("Warehouse sales", query.Name);
        Assert.Equal(7U, query.ConnectionId);
        Assert.True(query.Headers);
        Assert.False(query.RowNumbers);
        Assert.True(query.BackgroundRefresh);
        Assert.False(query.RefreshOnLoad);
        Assert.True(query.PreserveFormatting);
        Assert.Equal("insertClear", query.GrowShrinkType);
        Assert.Equal("xl/queryTables/queryTable1.xml", query.Source.QueryPartPath);
        Assert.Equal("rIdQueryTable", query.Source.RelationshipId);
        Assert.Equal("xl/connections.xml", query.Source.ConnectionPartPath);
        Assert.True(query.Source.Editable);
        var refresh = Assert.IsType<SpreadsheetTableQueryRefreshArtifact>(query.Refresh);
        Assert.True(refresh.PreserveSortFilterLayout);
        Assert.False(refresh.FieldIdWrapped);
        Assert.True(refresh.HeadersInLastRefresh);
        Assert.Equal(0U, refresh.MinimumVersion);
        Assert.Equal(3U, refresh.NextId);
        Assert.Equal(0U, refresh.UnboundColumnsLeft);
        Assert.Equal(0U, refresh.UnboundColumnsRight);
        Assert.Collection(refresh.Fields,
            field =>
            {
                Assert.Equal(1U, field.Id);
                Assert.Equal("Region", field.Name);
                Assert.True(field.DataBound);
                Assert.False(field.FillFormulas);
                Assert.False(field.Clipped);
                Assert.Equal(1U, field.TableColumnId);
            },
            field =>
            {
                Assert.Equal(2U, field.Id);
                Assert.Equal("Revenue", field.Name);
                Assert.True(field.DataBound);
                Assert.Equal(2U, field.TableColumnId);
            });
        Assert.Equal(["Legacy Region", "Legacy Revenue"], refresh.DeletedFieldNames);
        var refreshSort = Assert.IsType<SpreadsheetTableSortStateArtifact>(refresh.SortState);
        Assert.Equal("A2:B3", refreshSort.Reference);
        Assert.True(refreshSort.CaseSensitive);
        Assert.Collection(refreshSort.Conditions,
            condition =>
            {
                Assert.Equal("B2:B3", condition.Reference);
                Assert.True(condition.Descending);
                Assert.Null(condition.Icon);
            },
            condition =>
            {
                Assert.Equal("A2:A3", condition.Reference);
                Assert.Equal("3Arrows", condition.Icon.IconSet);
                Assert.Equal(0U, condition.Icon.IconId);
            });

        table.Name = "WarehouseTable";
        query.Name = "Warehouse sales refreshed";
        query.BackgroundRefresh = false;
        query.RefreshOnLoad = true;
        query.AutoFormatId = 3;
        query.ApplyFontFormats = true;
        refresh.PreserveSortFilterLayout = false;
        refresh.HeadersInLastRefresh = false;
        refresh.MinimumVersion = 1;
        refresh.Fields[0].Name = "Territory";
        refresh.Fields[0].DataBound = false;
        refresh.Fields[0].FillFormulas = true;
        refresh.Fields[1].Clipped = true;
        refresh.DeletedFieldNames[0] = "Legacy Territory";
        refreshSort.CaseSensitive = false;
        refreshSort.Conditions[0].Descending = false;
        refreshSort.Conditions[1].Icon.IconId = 1;
        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var output = exported.File.ToByteArray();
        AssertOffice2021Valid(output);
        Assert.Equal(connectionXml, ReadEntry(output, "xl/connections.xml"));
        Assert.Equal(relationshipXml, ReadEntry(output, "xl/tables/_rels/table1.xml.rels"));
        var queryXml = System.Text.Encoding.UTF8.GetString(ReadEntry(output, "xl/queryTables/queryTable1.xml"));
        Assert.Contains("name=\"Warehouse sales refreshed\"", queryXml);
        Assert.Contains("backgroundRefresh=\"0\"", queryXml);
        Assert.Contains("refreshOnLoad=\"1\"", queryXml);
        Assert.Contains("autoFormatId=\"3\"", queryXml);
        Assert.Contains("applyFontFormats=\"1\"", queryXml);
        Assert.Contains("preserveSortFilterLayout=\"0\"", queryXml);
        Assert.Contains("headersInLastRefresh=\"0\"", queryXml);
        Assert.Contains("minimumVersion=\"1\"", queryXml);
        Assert.Contains("id=\"1\" name=\"Territory\" dataBound=\"0\" tableColumnId=\"1\" fillFormulas=\"1\" clipped=\"0\"", queryXml);
        Assert.Contains("id=\"2\" name=\"Revenue\" dataBound=\"1\" tableColumnId=\"2\" clipped=\"1\"", queryXml);
        Assert.Contains("<x:queryTableFields count=\"2\">", queryXml);
        Assert.Contains("<x:deletedField name=\"Legacy Territory\"", queryXml);
        Assert.Contains("<x:deletedField name=\"Legacy Revenue\"", queryXml);
        Assert.Contains("<x:sortState ref=\"A2:B3\">", queryXml);
        Assert.Contains("<x:sortCondition ref=\"B2:B3\"", queryXml);
        Assert.Contains("<x:sortCondition ref=\"A2:A3\" sortBy=\"icon\" iconSet=\"3Arrows\" iconId=\"1\"", queryXml);
        Assert.Contains("<fixture:fieldOpaque value=\"kept\"", queryXml);
        Assert.Contains("<fixture:sortOpaque value=\"kept\"", queryXml);
        Assert.Contains("<fixture:opaque value=\"kept\"", queryXml);

        var reimported = Import(output);
        var edited = Assert.Single(reimported.Artifact.Workbook.Worksheets[0].Tables);
        Assert.Equal("WarehouseTable", edited.Name);
        Assert.Equal("Warehouse sales refreshed", edited.QueryTable.Name);
        Assert.False(edited.QueryTable.BackgroundRefresh);
        Assert.True(edited.QueryTable.RefreshOnLoad);
        Assert.Equal(3U, edited.QueryTable.AutoFormatId);
        Assert.True(edited.QueryTable.ApplyFontFormats);
        Assert.False(edited.QueryTable.Refresh.PreserveSortFilterLayout);
        Assert.False(edited.QueryTable.Refresh.HeadersInLastRefresh);
        Assert.Equal(1U, edited.QueryTable.Refresh.MinimumVersion);
        Assert.Equal("Territory", edited.QueryTable.Refresh.Fields[0].Name);
        Assert.False(edited.QueryTable.Refresh.Fields[0].DataBound);
        Assert.True(edited.QueryTable.Refresh.Fields[0].FillFormulas);
        Assert.True(edited.QueryTable.Refresh.Fields[1].DataBound);
        Assert.True(edited.QueryTable.Refresh.Fields[1].Clipped);
        Assert.Equal(["Legacy Territory", "Legacy Revenue"], edited.QueryTable.Refresh.DeletedFieldNames);
        Assert.False(edited.QueryTable.Refresh.SortState.CaseSensitive);
        Assert.False(edited.QueryTable.Refresh.SortState.Conditions[0].Descending);
        Assert.Equal(1U, edited.QueryTable.Refresh.SortState.Conditions[1].Icon.IconId);
        Assert.Equal(query.Source.QueryPartPath, edited.QueryTable.Source.QueryPartPath);
        Assert.Equal(query.Source.RelationshipId, edited.QueryTable.Source.RelationshipId);
    }

    [Fact]
    public void WorksheetQueryRefreshRejectsIdentityTopologyAndInvalidMetadata()
    {
        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(TableExportRequest().ToByteArray()));
        var source = AddQueryTableGraph(authored.File.ToByteArray());

        var imported = Import(source);
        imported.Artifact.Workbook.Worksheets[0].Tables[0].QueryTable.Refresh.Fields[0].Id = 99;
        var response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);

        imported = Import(source);
        imported.Artifact.Workbook.Worksheets[0].Tables[0].QueryTable.Refresh.Fields.RemoveAt(0);
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);

        imported = Import(source);
        imported.Artifact.Workbook.Worksheets[0].Tables[0].QueryTable.Refresh.NextId = 2;
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);

        imported = Import(source);
        imported.Artifact.Workbook.Worksheets[0].Tables[0].QueryTable.Refresh.MinimumVersion = 256;
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);

        imported = Import(source);
        imported.Artifact.Workbook.Worksheets[0].Tables[0].QueryTable.Refresh.Fields[0].FillFormulas = true;
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Contains("explicitly unbound", Assert.Single(response.Diagnostics).Message);

        imported = Import(source);
        imported.Artifact.Workbook.Worksheets[0].Tables[0].QueryTable.Refresh.Fields[0].DataBound = false;
        imported.Artifact.Workbook.Worksheets[0].Tables[0].QueryTable.Refresh.Fields[0].Clipped = true;
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Contains("explicitly bound", Assert.Single(response.Diagnostics).Message);

        imported = Import(source);
        imported.Artifact.Workbook.Worksheets[0].Tables[0].QueryTable.Refresh.DeletedFieldNames.RemoveAt(0);
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Contains("add or remove query refresh deleted fields", Assert.Single(response.Diagnostics).Message);

        imported = Import(source);
        imported.Artifact.Workbook.Worksheets[0].Tables[0].QueryTable.Refresh.DeletedFieldNames[1] = "Legacy Region";
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Contains("duplicate deleted-field name", Assert.Single(response.Diagnostics).Message);

        imported = Import(source);
        imported.Artifact.Workbook.Worksheets[0].Tables[0].QueryTable.Refresh.SortState.Conditions.RemoveAt(0);
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Contains("add or remove query refresh sort conditions", Assert.Single(response.Diagnostics).Message);

        imported = Import(source);
        imported.Artifact.Workbook.Worksheets[0].Tables[0].QueryTable.Refresh.SortState.Reference = "A2:C3";
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Contains("contained in the source table range", Assert.Single(response.Diagnostics).Message);
    }

    [Fact]
    public void OpaqueWorksheetQueryRefreshHistoryStaysHiddenAndPreserved()
    {
        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(TableExportRequest().ToByteArray()));
        var source = AddQueryTableGraph(authored.File.ToByteArray(), opaqueDeletedFields: true, opaqueSort: true);
        AssertOffice2021Valid(source);
        var imported = Import(source);
        var query = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Tables).QueryTable;
        Assert.NotNull(query.Refresh);
        Assert.Empty(query.Refresh.DeletedFieldNames);
        Assert.Null(query.Refresh.SortState);
        Assert.Equal("Region", query.Refresh.Fields[0].Name);
        query.Name = "Opaque history retained";
        query.Refresh.Fields[0].Name = "Territory";

        var exported = Export(imported.Artifact);
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(exported.File.ToByteArray());
        var queryXml = System.Text.Encoding.UTF8.GetString(ReadEntry(exported.File.ToByteArray(), "xl/queryTables/queryTable1.xml"));
        Assert.Contains("name=\"Opaque history retained\"", queryXml);
        Assert.Contains("name=\"Territory\"", queryXml);
        Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(queryXml, "<x:deletedField name=\"Legacy Region\"").Count);
        Assert.Contains("sortMethod=\"stroke\"", queryXml);
        Assert.Contains("<fixture:sortOpaque value=\"kept\"", queryXml);

        imported = Import(source);
        imported.Artifact.Workbook.Worksheets[0].Tables[0].QueryTable.Refresh.DeletedFieldNames.Add("Fabricated");
        var rejected = Export(imported.Artifact);
        Assert.False(rejected.Ok);
        Assert.Contains("absent/opaque query refresh deleted fields", Assert.Single(rejected.Diagnostics).Message);

        imported = Import(source);
        imported.Artifact.Workbook.Worksheets[0].Tables[0].QueryTable.Refresh.SortState = new SpreadsheetTableSortStateArtifact
        {
            Reference = "A2:B3",
        };
        rejected = Export(imported.Artifact);
        Assert.False(rejected.Ok);
        Assert.Contains("absent/opaque query refresh sort state", Assert.Single(rejected.Diagnostics).Message);
    }

    [Fact]
    public void OpaqueWorksheetQueryRefreshStaysHiddenDuringRootEdit()
    {
        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(TableExportRequest().ToByteArray()));
        var imported = Import(AddQueryTableGraph(authored.File.ToByteArray(), opaqueRefresh: true));
        var table = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Tables);
        Assert.NotNull(table.QueryTable);
        Assert.Null(table.QueryTable.Refresh);
        table.QueryTable.Name = "Opaque refresh retained";
        var response = Export(imported.Artifact);
        Assert.True(response.Ok, string.Join("\n", response.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var queryXml = System.Text.Encoding.UTF8.GetString(ReadEntry(response.File.ToByteArray(), "xl/queryTables/queryTable1.xml"));
        Assert.Contains("name=\"Opaque refresh retained\"", queryXml);
        Assert.Contains("tableColumnId=\"999\"", queryXml);

        imported = Import(AddQueryTableGraph(authored.File.ToByteArray(), opaqueRefresh: true));
        imported.Artifact.Workbook.Worksheets[0].Tables[0].QueryTable.Refresh = new SpreadsheetTableQueryRefreshArtifact();
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);
    }

    [Fact]
    public void WorksheetQueryTableRejectsFabricationRebindingAndBindingTamper()
    {
        var sourceFree = TableExportRequest();
        sourceFree.Artifact.Workbook.Worksheets[0].Tables[0].QueryTable = new SpreadsheetTableQueryArtifact
        {
            Name = "Fabricated query",
            ConnectionId = 1,
        };
        var response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(sourceFree.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);

        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(TableExportRequest().ToByteArray()));
        var imported = Import(AddQueryTableGraph(authored.File.ToByteArray()));
        imported.Artifact.Workbook.Worksheets[0].Tables[0].QueryTable.ConnectionId = 999;
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);

        imported = Import(AddQueryTableGraph(authored.File.ToByteArray()));
        imported.Artifact.Workbook.Worksheets[0].Tables[0].QueryTable.Source.ConnectionXmlSha256 = new string('0', 64);
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);
    }

    [Fact]
    public void UnsupportedWorksheetQueryTableRemainsByteExactAndHidden()
    {
        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(TableExportRequest().ToByteArray()));
        var source = AddQueryTableGraph(authored.File.ToByteArray(), addUnsupportedRelationship: true);
        var imported = Import(source);
        var table = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Tables);
        Assert.False(table.Source.Editable);
        Assert.Empty(table.Name);
        Assert.Null(table.QueryTable);
        var preserved = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.True(preserved.Ok, string.Join("\n", preserved.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal(ReadEntry(source, "xl/tables/table1.xml"), ReadEntry(preserved.File.ToByteArray(), "xl/tables/table1.xml"));
        Assert.Equal(ReadEntry(source, "xl/queryTables/queryTable1.xml"), ReadEntry(preserved.File.ToByteArray(), "xl/queryTables/queryTable1.xml"));
        Assert.Equal(ReadEntry(source, "xl/connections.xml"), ReadEntry(preserved.File.ToByteArray(), "xl/connections.xml"));
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
    public void ProtocolAuthorsAndImportsWorksheetTableColumnFormulasAndTotals()
    {
        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(FormulaTableExportRequest().ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var xml = System.Text.Encoding.UTF8.GetString(ReadEntry(exported.File.ToByteArray(), "xl/tables/table1.xml"));
        Assert.Contains("totalsRowLabel=\"Total\"", xml);
        Assert.Contains("totalsRowFunction=\"average\"", xml);
        Assert.Contains("<x:calculatedColumnFormula>[@Units]*2</x:calculatedColumnFormula>", xml);
        Assert.Contains("totalsRowFunction=\"custom\"", xml);
        Assert.Contains("<x:totalsRowFormula>SUBTOTAL(109,[Revenue])</x:totalsRowFormula>", xml);
        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));

        var imported = Import(exported.File.ToByteArray());
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var table = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Tables);
        Assert.True(table.ShowTotals);
        Assert.Equal(3, table.Columns.Count);
        Assert.Equal("Total", table.Columns[0].TotalsRowLabel);
        Assert.Equal("average", table.Columns[1].TotalsRowFunction);
        Assert.Equal("=[@Units]*2", table.Columns[2].CalculatedColumnFormula);
        Assert.Equal("custom", table.Columns[2].TotalsRowFunction);
        Assert.Equal("=SUBTOTAL(109,[Revenue])", table.Columns[2].TotalsRowFormula);
    }

    [Fact]
    public void SourcePreservingWorksheetTableColumnFormulaEditKeepsPartIdentity()
    {
        var first = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(FormulaTableExportRequest().ToByteArray()));
        var imported = Import(first.File.ToByteArray());
        var table = imported.Artifact.Workbook.Worksheets[0].Tables[0];
        var path = table.Source.TablePartPath;
        var relationshipId = table.Source.RelationshipId;
        table.Columns[1].TotalsRowFunction = "max";
        table.Columns[2].CalculatedColumnFormula = "=[@Units]*3";
        table.Columns[2].TotalsRowFormula = "=SUBTOTAL(109,[Revenue])+1";
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
        Assert.Equal("max", edited.Columns[1].TotalsRowFunction);
        Assert.Equal("=[@Units]*3", edited.Columns[2].CalculatedColumnFormula);
        Assert.Equal("=SUBTOTAL(109,[Revenue])+1", edited.Columns[2].TotalsRowFormula);
        Assert.Equal(path, edited.Source.TablePartPath);
        Assert.Equal(relationshipId, edited.Source.RelationshipId);
    }

    [Fact]
    public void ProtocolRejectsInvalidWorksheetTableColumnFormulaProfiles()
    {
        var request = FormulaTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].Columns[2].CalculatedColumnFormula = "[@Units]*2";
        var response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);

        request = FormulaTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].Columns[2].TotalsRowFunction = "sum";
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);

        request = FormulaTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].ShowTotals = false;
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);
    }

    [Fact]
    public void ProtocolAuthorsAndImportsWorksheetTableValueAndCustomFilters()
    {
        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(FilterTableExportRequest().ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var xml = System.Text.Encoding.UTF8.GetString(ReadEntry(exported.File.ToByteArray(), "xl/tables/table1.xml"));
        Assert.Contains("<x:filterColumn colId=\"0\"><x:filters blank=\"1\"><x:filter val=\"North\" /><x:filter val=\"South\" /></x:filters></x:filterColumn>", xml);
        Assert.Contains("<x:filterColumn colId=\"1\"><x:customFilters and=\"1\"><x:customFilter operator=\"greaterThanOrEqual\" val=\"80\" /><x:customFilter operator=\"lessThanOrEqual\" val=\"120\" /></x:customFilters></x:filterColumn>", xml);
        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));

        var imported = Import(exported.File.ToByteArray());
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var filters = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Tables).Filters;
        Assert.Equal(2, filters.Count);
        Assert.Equal(0U, filters[0].ColumnIndex);
        Assert.Equal(["North", "South"], filters[0].Values.Values);
        Assert.True(filters[0].Values.IncludeBlank);
        Assert.Equal(1U, filters[1].ColumnIndex);
        Assert.True(filters[1].Custom.MatchAll);
        Assert.Equal("greaterThanOrEqual", filters[1].Custom.Criteria[0].Operator);
        Assert.Equal("80", filters[1].Custom.Criteria[0].Value);
    }

    [Fact]
    public void SourcePreservingWorksheetTableFilterEditKeepsPartIdentity()
    {
        var first = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(FilterTableExportRequest().ToByteArray()));
        var imported = Import(first.File.ToByteArray());
        var table = imported.Artifact.Workbook.Worksheets[0].Tables[0];
        var path = table.Source.TablePartPath;
        var relationshipId = table.Source.RelationshipId;
        table.Filters[0].Values.Values.Clear();
        table.Filters[0].Values.Values.Add("North");
        table.Filters[0].Values.IncludeBlank = false;
        table.Filters[1].Custom.MatchAll = false;
        table.Filters[1].Custom.Criteria[0].Value = "100";
        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var edited = Assert.Single(Import(exported.File.ToByteArray()).Artifact.Workbook.Worksheets[0].Tables);
        Assert.Equal(["North"], edited.Filters[0].Values.Values);
        Assert.False(edited.Filters[0].Values.IncludeBlank);
        Assert.False(edited.Filters[1].Custom.MatchAll);
        Assert.Equal("100", edited.Filters[1].Custom.Criteria[0].Value);
        Assert.Equal(path, edited.Source.TablePartPath);
        Assert.Equal(relationshipId, edited.Source.RelationshipId);
    }

    [Fact]
    public void ProtocolRejectsInvalidWorksheetTableFilterProfiles()
    {
        var request = FilterTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].Filters[0].ColumnIndex = 2;
        var response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);

        request = FilterTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].Filters[1].ColumnIndex = 0;
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);

        request = FilterTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].Filters[1].Custom.Criteria.Add(new SpreadsheetTableCustomFilterCriterionArtifact { Operator = "equal", Value = "100" });
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);

        request = FilterTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].ShowFilterButton = false;
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);
    }

    [Fact]
    public void ProtocolAuthorsAndImportsWorksheetTableDateDynamicAndTop10Filters()
    {
        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(AdvancedFilterTableExportRequest().ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var xml = System.Text.Encoding.UTF8.GetString(ReadEntry(exported.File.ToByteArray(), "xl/tables/table1.xml"));
        Assert.Contains("calendarType=\"gregorian\"", xml);
        Assert.Contains("<x:dateGroupItem year=\"2026\" dateTimeGrouping=\"day\" month=\"7\" day=\"15\" />", xml);
        Assert.Contains("<x:dynamicFilter type=\"today\" val=\"45853\" maxVal=\"45854\" />", xml);
        Assert.Contains("<x:top10 top=\"1\" percent=\"1\" val=\"10\" filterVal=\"95\" />", xml);
        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));

        var imported = Import(exported.File.ToByteArray());
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var filters = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Tables).Filters;
        Assert.Equal(3, filters.Count);
        Assert.Equal("gregorian", filters[0].Values.CalendarType);
        Assert.Empty(filters[0].Values.Values);
        var group = Assert.Single(filters[0].Values.DateGroups);
        Assert.Equal("day", group.Grouping);
        Assert.Equal(2026U, group.Year);
        Assert.Equal(7U, group.Month);
        Assert.Equal(15U, group.Day);
        Assert.Equal("today", filters[1].Dynamic.Type);
        Assert.Equal(45853, filters[1].Dynamic.Value);
        Assert.Equal(45854, filters[1].Dynamic.MaxValue);
        Assert.True(filters[2].Top10.Top);
        Assert.True(filters[2].Top10.Percent);
        Assert.Equal(10, filters[2].Top10.Value);
        Assert.Equal(95, filters[2].Top10.FilterValue);
    }

    [Fact]
    public void SourcePreservingAdvancedWorksheetTableFilterEditKeepsPartIdentity()
    {
        var first = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(AdvancedFilterTableExportRequest().ToByteArray()));
        var imported = Import(first.File.ToByteArray());
        var table = imported.Artifact.Workbook.Worksheets[0].Tables[0];
        var path = table.Source.TablePartPath;
        var relationshipId = table.Source.RelationshipId;
        table.Filters[0].Values.DateGroups[0].Day = 16;
        table.Filters[1].Dynamic.Type = "yesterday";
        table.Filters[1].Dynamic.Value = 45852;
        table.Filters[1].Dynamic.MaxValue = 45853;
        table.Filters[2].Top10.Top = false;
        table.Filters[2].Top10.Percent = false;
        table.Filters[2].Top10.Value = 5;
        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var edited = Assert.Single(Import(exported.File.ToByteArray()).Artifact.Workbook.Worksheets[0].Tables);
        Assert.Equal(16U, edited.Filters[0].Values.DateGroups[0].Day);
        Assert.Equal("yesterday", edited.Filters[1].Dynamic.Type);
        Assert.False(edited.Filters[2].Top10.Top);
        Assert.False(edited.Filters[2].Top10.Percent);
        Assert.Equal(5, edited.Filters[2].Top10.Value);
        Assert.Equal(path, edited.Source.TablePartPath);
        Assert.Equal(relationshipId, edited.Source.RelationshipId);
    }

    [Fact]
    public void ProtocolRejectsInvalidAdvancedWorksheetTableFilterProfiles()
    {
        var request = AdvancedFilterTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].Filters[0].Values.DateGroups[0] = new SpreadsheetTableDateGroupItemArtifact
            { Year = 2026, Day = 15, Grouping = "day" };
        var response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);

        request = AdvancedFilterTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].Filters[1].Dynamic.Type = "thisDecade";
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);

        request = AdvancedFilterTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].Filters[2].Top10.Value = 101;
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);
    }

    [Fact]
    public void ProtocolAuthorsAndImportsWorksheetTableIconFiltersAndSorts()
    {
        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(IconTableExportRequest().ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var xml = System.Text.Encoding.UTF8.GetString(ReadEntry(exported.File.ToByteArray(), "xl/tables/table1.xml"));
        Assert.Contains("<x:iconFilter iconSet=\"3Arrows\" iconId=\"0\" />", xml);
        Assert.Contains("<x:iconFilter iconSet=\"3Flags\" />", xml);
        Assert.Contains("<x:sortCondition ref=\"B2:B3\" descending=\"1\" sortBy=\"icon\" iconSet=\"5Rating\" iconId=\"4\" />", xml);
        Assert.Contains("<x:sortCondition ref=\"A2:A3\" sortBy=\"icon\" iconSet=\"3Symbols2\" />", xml);
        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));

        var imported = Import(exported.File.ToByteArray());
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var table = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Tables);
        Assert.True(table.Source.Editable);
        Assert.Equal(2, table.Filters.Count);
        Assert.Equal(SpreadsheetTableFilterArtifact.CriteriaOneofCase.Icon, table.Filters[0].CriteriaCase);
        Assert.Equal("3Arrows", table.Filters[0].Icon.IconSet);
        Assert.True(table.Filters[0].Icon.HasIconId);
        Assert.Equal(0U, table.Filters[0].Icon.IconId);
        Assert.Equal("3Flags", table.Filters[1].Icon.IconSet);
        Assert.False(table.Filters[1].Icon.HasIconId);
        Assert.Equal("5Rating", table.SortState.Conditions[0].Icon.IconSet);
        Assert.Equal(4U, table.SortState.Conditions[0].Icon.IconId);
        Assert.Equal("3Symbols2", table.SortState.Conditions[1].Icon.IconSet);
        Assert.False(table.SortState.Conditions[1].Icon.HasIconId);
    }

    [Fact]
    public void SourcePreservingWorksheetTableIconEditKeepsPartIdentity()
    {
        var first = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(IconTableExportRequest().ToByteArray()));
        var imported = Import(first.File.ToByteArray());
        var table = imported.Artifact.Workbook.Worksheets[0].Tables[0];
        var path = table.Source.TablePartPath;
        var relationshipId = table.Source.RelationshipId;
        table.Filters[0].Icon.IconId = 2;
        table.Filters[1].Icon.IconId = 1;
        table.SortState.Conditions[0].Icon.IconSet = "4Rating";
        table.SortState.Conditions[0].Icon.IconId = 3;
        table.SortState.Conditions[1].Icon.IconId = 2;
        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var edited = Assert.Single(Import(exported.File.ToByteArray()).Artifact.Workbook.Worksheets[0].Tables);
        Assert.Equal(2U, edited.Filters[0].Icon.IconId);
        Assert.Equal(1U, edited.Filters[1].Icon.IconId);
        Assert.Equal("4Rating", edited.SortState.Conditions[0].Icon.IconSet);
        Assert.Equal(3U, edited.SortState.Conditions[0].Icon.IconId);
        Assert.Equal(2U, edited.SortState.Conditions[1].Icon.IconId);
        Assert.Equal(path, edited.Source.TablePartPath);
        Assert.Equal(relationshipId, edited.Source.RelationshipId);
    }

    [Fact]
    public void ProtocolRejectsInvalidWorksheetTableIconProfiles()
    {
        var request = IconTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].Filters[0].Icon.IconSet = "3Stars";
        var response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);

        request = IconTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].Filters[0].Icon.IconId = 3;
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);

        request = IconTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].SortState.Conditions[0].Icon.IconId = 5;
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);
    }

    [Fact]
    public void ProtocolAuthorsAndImportsWorksheetTableColorFiltersAndSorts()
    {
        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(ColorTableExportRequest().ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var bytes = exported.File.ToByteArray();
        var tableXml = System.Text.Encoding.UTF8.GetString(ReadEntry(bytes, "xl/tables/table1.xml"));
        Assert.Contains("<x:colorFilter dxfId=\"0\" cellColor=\"1\" />", tableXml);
        Assert.Contains("<x:colorFilter dxfId=\"1\" cellColor=\"0\" />", tableXml);
        Assert.Contains("<x:sortCondition ref=\"B2:B3\" descending=\"1\" sortBy=\"fontColor\" dxfId=\"1\" />", tableXml);
        Assert.Contains("<x:sortCondition ref=\"A2:A3\" sortBy=\"cellColor\" dxfId=\"0\" />", tableXml);
        using (var stream = new MemoryStream(bytes))
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
            var formats = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.DifferentialFormats!.Elements<DifferentialFormat>().ToArray();
            Assert.Equal(2, formats.Length);
            Assert.Equal("FFE11D48", formats[0].Fill!.PatternFill!.ForegroundColor!.Rgb!.Value);
            Assert.Equal(4U, formats[1].Font!.Color!.Theme!.Value);
            Assert.Equal(-0.25D, formats[1].Font!.Color!.Tint!.Value);
        }

        var imported = Import(bytes);
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var table = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Tables);
        Assert.True(table.Source.Editable);
        Assert.Equal(SpreadsheetTableColorArtifact.TargetOneofCase.CellColor, table.Filters[0].Color.TargetCase);
        Assert.Equal("E11D48", table.Filters[0].Color.Color.Rgb);
        Assert.Equal(SpreadsheetTableColorArtifact.TargetOneofCase.FontColor, table.Filters[1].Color.TargetCase);
        Assert.Equal(4U, table.Filters[1].Color.Color.Theme);
        Assert.Equal(-0.25D, table.Filters[1].Color.Color.Tint);
        Assert.Equal(SpreadsheetTableColorArtifact.TargetOneofCase.FontColor, table.SortState.Conditions[0].Color.TargetCase);
        Assert.Equal(SpreadsheetTableColorArtifact.TargetOneofCase.CellColor, table.SortState.Conditions[1].Color.TargetCase);
    }

    [Fact]
    public void SourcePreservingWorksheetTableColorEditAppendsDifferentialFormatsAndKeepsPartIdentity()
    {
        var first = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(ColorTableExportRequest().ToByteArray()));
        var originalFormats = DifferentialFormatXml(first.File.ToByteArray());
        var imported = Import(first.File.ToByteArray());
        var table = imported.Artifact.Workbook.Worksheets[0].Tables[0];
        var path = table.Source.TablePartPath;
        var relationshipId = table.Source.RelationshipId;
        table.Filters[0].Color.Color.Rgb = "22C55E";
        table.SortState.Conditions[1].Color.Color.Rgb = "22C55E";
        table.Filters[1].Color.Color.Rgb = "2563EB";
        table.SortState.Conditions[0].Color.Color.Rgb = "2563EB";

        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var formats = DifferentialFormatXml(exported.File.ToByteArray());
        Assert.Equal(4, formats.Length);
        Assert.Equal(originalFormats, formats[..originalFormats.Length]);
        var edited = Assert.Single(Import(exported.File.ToByteArray()).Artifact.Workbook.Worksheets[0].Tables);
        Assert.Equal("22C55E", edited.Filters[0].Color.Color.Rgb);
        Assert.Equal("2563EB", edited.Filters[1].Color.Color.Rgb);
        Assert.Equal(path, edited.Source.TablePartPath);
        Assert.Equal(relationshipId, edited.Source.RelationshipId);
    }

    [Fact]
    public void ProtocolRejectsInvalidWorksheetTableColorProfiles()
    {
        var request = ColorTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].Filters[0].Color.ClearTarget();
        var response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);

        request = ColorTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].Filters[0].Color.CellColor = false;
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);

        request = ColorTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].Filters[0].Color.Color.Rgb = "xyz";
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);

        request = ColorTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].SortState.Conditions[0].Icon = new SpreadsheetTableIconArtifact { IconSet = "3Arrows" };
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);
    }

    [Fact]
    public void UnsupportedColorWorksheetTableFilterRemainsByteExactAndReadOnly()
    {
        var request = TableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Cells[3].NumberFormatCode = "0";
        var first = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        var bytes = MutateTableWithColorFilter(first.File.ToByteArray());
        var imported = Import(bytes);
        var table = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Tables);
        Assert.False(table.Source.Editable);
        Assert.Empty(table.Filters);
        var preserved = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.True(preserved.Ok, string.Join("\n", preserved.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal(ReadEntry(bytes, "xl/tables/table1.xml"), ReadEntry(preserved.File.ToByteArray(), "xl/tables/table1.xml"));
    }

    [Fact]
    public void ProtocolAuthorsAndImportsWorksheetTableValueSortState()
    {
        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(SortTableExportRequest().ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var xml = System.Text.Encoding.UTF8.GetString(ReadEntry(exported.File.ToByteArray(), "xl/tables/table1.xml"));
        Assert.Contains("<x:sortState ref=\"A2:B3\" caseSensitive=\"1\"><x:sortCondition ref=\"B2:B3\" descending=\"1\" /><x:sortCondition ref=\"A2:A3\" /></x:sortState>", xml);
        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));

        var imported = Import(exported.File.ToByteArray());
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var sort = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Tables).SortState;
        Assert.Equal("A2:B3", sort.Reference);
        Assert.True(sort.CaseSensitive);
        Assert.Equal(2, sort.Conditions.Count);
        Assert.Equal("B2:B3", sort.Conditions[0].Reference);
        Assert.True(sort.Conditions[0].Descending);
        Assert.Equal("A2:A3", sort.Conditions[1].Reference);
        Assert.False(sort.Conditions[1].Descending);
    }

    [Fact]
    public void SourcePreservingWorksheetTableSortEditKeepsPartIdentity()
    {
        var first = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(SortTableExportRequest().ToByteArray()));
        var imported = Import(first.File.ToByteArray());
        var table = imported.Artifact.Workbook.Worksheets[0].Tables[0];
        var path = table.Source.TablePartPath;
        var relationshipId = table.Source.RelationshipId;
        table.SortState.CaseSensitive = false;
        table.SortState.Conditions[0].Descending = false;
        table.SortState.Conditions[1].Descending = true;
        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var edited = Assert.Single(Import(exported.File.ToByteArray()).Artifact.Workbook.Worksheets[0].Tables);
        Assert.False(edited.SortState.CaseSensitive);
        Assert.False(edited.SortState.Conditions[0].Descending);
        Assert.True(edited.SortState.Conditions[1].Descending);
        Assert.Equal(path, edited.Source.TablePartPath);
        Assert.Equal(relationshipId, edited.Source.RelationshipId);
    }

    [Fact]
    public void ProtocolRejectsInvalidWorksheetTableSortProfiles()
    {
        var request = SortTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].SortState.Reference = "B2:C3";
        var response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);

        request = SortTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].SortState.Conditions[0].Reference = "A2:B3";
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);

        request = SortTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].SortState.Conditions[1].Reference = "B2:B3";
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);

        request = SortTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].ShowFilterButton = false;
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_table", Assert.Single(response.Diagnostics).Code);
    }

    [Fact]
    public void UnsupportedColorWorksheetTableSortRemainsByteExactAndReadOnly()
    {
        var first = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(SortTableExportRequest().ToByteArray()));
        var bytes = MutateTableWithColorSort(first.File.ToByteArray());
        var imported = Import(bytes);
        var table = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Tables);
        Assert.False(table.Source.Editable);
        Assert.Null(table.SortState);
        var preserved = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportXlsx,
            Family = ArtifactFamily.Workbook,
            Artifact = imported.Artifact,
        }.ToByteArray()));
        Assert.True(preserved.Ok, string.Join("\n", preserved.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal(ReadEntry(bytes, "xl/tables/table1.xml"), ReadEntry(preserved.File.ToByteArray(), "xl/tables/table1.xml"));
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

    private static CodecRequest FormulaTableExportRequest()
    {
        var request = ExportRequest();
        var sheet = request.Artifact.Workbook.Worksheets[0];
        sheet.Cells.Clear();
        sheet.Cells.Add(new CellArtifact { Row = 0, Column = 0, StringValue = "Product" });
        sheet.Cells.Add(new CellArtifact { Row = 0, Column = 1, StringValue = "Units" });
        sheet.Cells.Add(new CellArtifact { Row = 0, Column = 2, StringValue = "Revenue" });
        sheet.Cells.Add(new CellArtifact { Row = 1, Column = 0, StringValue = "North" });
        sheet.Cells.Add(new CellArtifact { Row = 1, Column = 1, NumberValue = 2 });
        sheet.Cells.Add(new CellArtifact { Row = 1, Column = 2, Formula = "=B2*2", NumberValue = 4 });
        sheet.Cells.Add(new CellArtifact { Row = 2, Column = 0, StringValue = "South" });
        sheet.Cells.Add(new CellArtifact { Row = 2, Column = 1, NumberValue = 3 });
        sheet.Cells.Add(new CellArtifact { Row = 2, Column = 2, Formula = "=B3*2", NumberValue = 6 });
        sheet.Cells.Add(new CellArtifact { Row = 3, Column = 0, StringValue = "Total" });
        sheet.Cells.Add(new CellArtifact { Row = 3, Column = 1, Formula = "=AVERAGE(B2:B3)", NumberValue = 2.5 });
        sheet.Cells.Add(new CellArtifact { Row = 3, Column = 2, Formula = "=SUBTOTAL(109,C2:C3)", NumberValue = 10 });
        var table = new SpreadsheetTableArtifact
        {
            Id = "table/formulas",
            Name = "FormulaTable",
            Reference = "A1:C4",
            HasHeaders = true,
            ShowTotals = true,
            ShowFilterButton = true,
            StyleName = "TableStyleMedium4",
            ShowRowStripes = true,
        };
        table.ColumnNames.Add("Product");
        table.ColumnNames.Add("Units");
        table.ColumnNames.Add("Revenue");
        table.Columns.Add(new SpreadsheetTableColumnArtifact { Name = "Product", TotalsRowFunction = "none", TotalsRowLabel = "Total" });
        table.Columns.Add(new SpreadsheetTableColumnArtifact { Name = "Units", TotalsRowFunction = "average" });
        table.Columns.Add(new SpreadsheetTableColumnArtifact
        {
            Name = "Revenue",
            CalculatedColumnFormula = "=[@Units]*2",
            TotalsRowFunction = "custom",
            TotalsRowFormula = "=SUBTOTAL(109,[Revenue])",
        });
        sheet.Tables.Add(table);
        return request;
    }

    private static CodecRequest FilterTableExportRequest()
    {
        var request = TableExportRequest();
        var table = request.Artifact.Workbook.Worksheets[0].Tables[0];
        var values = new SpreadsheetTableValueFilterArtifact { IncludeBlank = true };
        values.Values.Add("North");
        values.Values.Add("South");
        table.Filters.Add(new SpreadsheetTableFilterArtifact { ColumnIndex = 0, Values = values });
        var custom = new SpreadsheetTableCustomFilterArtifact { MatchAll = true };
        custom.Criteria.Add(new SpreadsheetTableCustomFilterCriterionArtifact { Operator = "greaterThanOrEqual", Value = "80" });
        custom.Criteria.Add(new SpreadsheetTableCustomFilterCriterionArtifact { Operator = "lessThanOrEqual", Value = "120" });
        table.Filters.Add(new SpreadsheetTableFilterArtifact { ColumnIndex = 1, Custom = custom });
        return request;
    }

    private static CodecRequest AdvancedFilterTableExportRequest()
    {
        var request = ExportRequest();
        var sheet = request.Artifact.Workbook.Worksheets[0];
        sheet.Cells.Clear();
        foreach (var (column, value) in new[] { "Date", "Status", "Score" }.Select((value, column) => (column, value)))
            sheet.Cells.Add(new CellArtifact { Row = 0, Column = (uint)column, StringValue = value });
        sheet.Cells.Add(new CellArtifact { Row = 1, Column = 0, NumberValue = 45853, NumberFormatCode = "yyyy-mm-dd" });
        sheet.Cells.Add(new CellArtifact { Row = 1, Column = 1, StringValue = "ready" });
        sheet.Cells.Add(new CellArtifact { Row = 1, Column = 2, NumberValue = 95 });
        sheet.Cells.Add(new CellArtifact { Row = 2, Column = 0, NumberValue = 45854, NumberFormatCode = "yyyy-mm-dd" });
        sheet.Cells.Add(new CellArtifact { Row = 2, Column = 1, StringValue = "pending" });
        sheet.Cells.Add(new CellArtifact { Row = 2, Column = 2, NumberValue = 80 });
        var table = new SpreadsheetTableArtifact
        {
            Id = "table/advanced-filters", Name = "AdvancedFilterTable", Reference = "A1:C3", HasHeaders = true,
            ShowFilterButton = true, StyleName = "TableStyleMedium4", ShowRowStripes = true,
        };
        table.ColumnNames.Add(["Date", "Status", "Score"]);
        var values = new SpreadsheetTableValueFilterArtifact { CalendarType = "gregorian" };
        values.DateGroups.Add(new SpreadsheetTableDateGroupItemArtifact { Year = 2026, Month = 7, Day = 15, Grouping = "day" });
        table.Filters.Add(new SpreadsheetTableFilterArtifact { ColumnIndex = 0, Values = values });
        table.Filters.Add(new SpreadsheetTableFilterArtifact
        {
            ColumnIndex = 1,
            Dynamic = new SpreadsheetTableDynamicFilterArtifact { Type = "today", Value = 45853, MaxValue = 45854 },
        });
        table.Filters.Add(new SpreadsheetTableFilterArtifact
        {
            ColumnIndex = 2,
            Top10 = new SpreadsheetTableTop10FilterArtifact { Top = true, Percent = true, Value = 10, FilterValue = 95 },
        });
        sheet.Tables.Add(table);
        return request;
    }

    private static CodecRequest SortTableExportRequest()
    {
        var request = FilterTableExportRequest();
        var sort = new SpreadsheetTableSortStateArtifact { Reference = "A2:B3", CaseSensitive = true };
        sort.Conditions.Add(new SpreadsheetTableSortConditionArtifact { Reference = "B2:B3", Descending = true });
        sort.Conditions.Add(new SpreadsheetTableSortConditionArtifact { Reference = "A2:A3" });
        request.Artifact.Workbook.Worksheets[0].Tables[0].SortState = sort;
        return request;
    }

    private static CodecRequest IconTableExportRequest()
    {
        var request = TableExportRequest();
        var table = request.Artifact.Workbook.Worksheets[0].Tables[0];
        table.Filters.Add(new SpreadsheetTableFilterArtifact
        {
            ColumnIndex = 0,
            Icon = new SpreadsheetTableIconArtifact { IconSet = "3Arrows", IconId = 0 },
        });
        table.Filters.Add(new SpreadsheetTableFilterArtifact
        {
            ColumnIndex = 1,
            Icon = new SpreadsheetTableIconArtifact { IconSet = "3Flags" },
        });
        var sort = new SpreadsheetTableSortStateArtifact { Reference = "A2:B3" };
        sort.Conditions.Add(new SpreadsheetTableSortConditionArtifact
        {
            Reference = "B2:B3", Descending = true,
            Icon = new SpreadsheetTableIconArtifact { IconSet = "5Rating", IconId = 4 },
        });
        sort.Conditions.Add(new SpreadsheetTableSortConditionArtifact
        {
            Reference = "A2:A3",
            Icon = new SpreadsheetTableIconArtifact { IconSet = "3Symbols2" },
        });
        table.SortState = sort;
        return request;
    }

    private static CodecRequest ColorTableExportRequest()
    {
        var request = TableExportRequest();
        var table = request.Artifact.Workbook.Worksheets[0].Tables[0];
        table.Filters.Add(new SpreadsheetTableFilterArtifact
        {
            ColumnIndex = 0,
            Color = CellTableColor(new SpreadsheetColor { Rgb = "E11D48" }),
        });
        table.Filters.Add(new SpreadsheetTableFilterArtifact
        {
            ColumnIndex = 1,
            Color = FontTableColor(new SpreadsheetColor { Theme = 4, Tint = -0.25D }),
        });
        var sort = new SpreadsheetTableSortStateArtifact { Reference = "A2:B3" };
        sort.Conditions.Add(new SpreadsheetTableSortConditionArtifact
        {
            Reference = "B2:B3", Descending = true,
            Color = FontTableColor(new SpreadsheetColor { Theme = 4, Tint = -0.25D }),
        });
        sort.Conditions.Add(new SpreadsheetTableSortConditionArtifact
        {
            Reference = "A2:A3",
            Color = CellTableColor(new SpreadsheetColor { Rgb = "E11D48" }),
        });
        table.SortState = sort;
        return request;
    }

    private static SpreadsheetTableColorArtifact CellTableColor(SpreadsheetColor color) => new() { CellColor = true, Color = color };
    private static SpreadsheetTableColorArtifact FontTableColor(SpreadsheetColor color) => new() { FontColor = true, Color = color };

    private static string[] DifferentialFormatXml(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        return document.WorkbookPart?.WorkbookStylesPart?.Stylesheet?.DifferentialFormats?.Elements<DifferentialFormat>().Select(item => item.OuterXml).ToArray() ?? [];
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

    private static byte[] MutateTableWithColorSort(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("xl/tables/table1.xml") ?? throw new InvalidOperationException("Worksheet table is missing.");
            XDocument table;
            using (var reader = new StreamReader(entry.Open())) table = XDocument.Parse(reader.ReadToEnd(), LoadOptions.PreserveWhitespace);
            var spreadsheet = table.Root!.Name.Namespace;
            var condition = table.Root.Element(spreadsheet + "autoFilter")!.Element(spreadsheet + "sortState")!.Elements(spreadsheet + "sortCondition").First();
            condition.SetAttributeValue("sortBy", "cellColor");
            condition.SetAttributeValue("dxfId", 0);
            entry.Delete();
            var replacement = archive.CreateEntry("xl/tables/table1.xml");
            using var writer = new StreamWriter(replacement.Open());
            writer.Write(table.ToString(SaveOptions.DisableFormatting));
        }
        return stream.ToArray();
    }

    private static byte[] MutateTableWithColorFilter(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var stylesheet = document.WorkbookPart?.WorkbookStylesPart?.Stylesheet
                ?? throw new InvalidOperationException("Color-filter fixture requires a stylesheet.");
            var differentialFormats = stylesheet.DifferentialFormats;
            if (differentialFormats is null)
            {
                differentialFormats = new DifferentialFormats();
                stylesheet.InsertBefore(differentialFormats, stylesheet.TableStyles);
            }
            differentialFormats.Append(new DifferentialFormat());
            differentialFormats.Count = checked((uint)differentialFormats.ChildElements.Count);
            stylesheet.Save();
        }
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("xl/tables/table1.xml") ?? throw new InvalidOperationException("Worksheet table is missing.");
            XDocument table;
            using (var reader = new StreamReader(entry.Open())) table = XDocument.Parse(reader.ReadToEnd(), LoadOptions.PreserveWhitespace);
            var spreadsheet = table.Root!.Name.Namespace;
            table.Root.Element(spreadsheet + "autoFilter")!.Add(new XElement(spreadsheet + "filterColumn", new XAttribute("colId", 0),
                new XElement(spreadsheet + "colorFilter", new XAttribute("dxfId", 0), new XAttribute("cellColor", 1))));
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

    private static CodecResponse Export(ArtifactEnvelope artifact) => CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(new CodecRequest
    {
        ProtocolVersion = CodecProtocol.ProtocolVersion,
        Operation = CodecOperation.ExportXlsx,
        Family = ArtifactFamily.Workbook,
        Artifact = artifact,
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

    private static byte[] AddQueryTableGraph(
        byte[] bytes,
        bool addUnsupportedRelationship = false,
        bool opaqueRefresh = false,
        bool opaqueDeletedFields = false,
        bool opaqueSort = false)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var workbookPart = document.WorkbookPart!;
            var connectionsPart = workbookPart.AddNewPart<ConnectionsPart>("rIdConnections");
            WritePart(connectionsPart, """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <x:connections xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <x:connection id="7" name="Fixture warehouse" type="5" refreshedVersion="8" background="1" refreshOnLoad="0" saveData="1">
                    <x:dbPr connection="Provider=Fixture.Provider;Data Source=fixture.invalid" command="SELECT Region, Revenue FROM Sales" commandType="2"/>
                  </x:connection>
                </x:connections>
                """);
            var tablePart = workbookPart.WorksheetParts.Single().TableDefinitionParts.Single();
            if (addUnsupportedRelationship)
                tablePart.AddExternalRelationship("urn:open-office-artifact-tool:unsupported-query-companion", new Uri("https://fixture.invalid/query"), "rIdUnsupportedQueryCompanion");
            var queryPart = tablePart.AddNewPart<QueryTablePart>("rIdQueryTable");
            var secondDeletedFieldName = opaqueDeletedFields ? "Legacy Region" : "Legacy Revenue";
            var sortOpaqueAttribute = opaqueSort ? " sortMethod=\"stroke\"" : string.Empty;
            WritePart(queryPart, $$"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <x:queryTable xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:fixture="urn:open-office-artifact-tool:query-fixture" name="Warehouse sales" headers="1" rowNumbers="0" disableRefresh="0" backgroundRefresh="1" firstBackgroundRefresh="0" refreshOnLoad="0" growShrinkType="insertClear" fillFormulas="0" removeDataOnSave="0" disableEdit="0" preserveFormatting="1" adjustColumnWidth="1" intermediate="0" connectionId="7">
                  <x:queryTableRefresh preserveSortFilterLayout="1" fieldIdWrapped="0" headersInLastRefresh="1" minimumVersion="0" nextId="3" unboundColumnsLeft="0" unboundColumnsRight="0">
                    <x:queryTableFields count="2">
                      <x:queryTableField id="1" name="Region" dataBound="1" tableColumnId="1" fillFormulas="0" clipped="0">
                        <x:extLst><x:ext uri="{71C44015-E485-449B-93BE-190C959F820F}"><fixture:fieldOpaque value="kept"/></x:ext></x:extLst>
                      </x:queryTableField>
                      <x:queryTableField id="2" name="Revenue" dataBound="1" tableColumnId="{{(opaqueRefresh ? 999 : 2)}}"/>
                    </x:queryTableFields>
                    <x:queryTableDeletedFields count="2">
                      <x:deletedField name="Legacy Region"/>
                      <x:deletedField name="{{secondDeletedFieldName}}"/>
                    </x:queryTableDeletedFields>
                    <x:sortState ref="A2:B3" caseSensitive="1"{{sortOpaqueAttribute}}>
                      <x:sortCondition ref="B2:B3" descending="1"/>
                      <x:sortCondition ref="A2:A3" sortBy="icon" iconSet="3Arrows" iconId="0"/>
                      <x:extLst><x:ext uri="{A1E10EA8-3B88-4BE3-9884-625AB42E9DDC}"><fixture:sortOpaque value="kept"/></x:ext></x:extLst>
                    </x:sortState>
                  </x:queryTableRefresh>
                  <x:extLst><x:ext uri="{A1D56E5F-35B8-4C51-9C80-779E6A39D52B}"><fixture:opaque value="kept"/></x:ext></x:extLst>
                </x:queryTable>
                """);
        }
        return stream.ToArray();
    }

    private static void WritePart(OpenXmlPart part, string xml)
    {
        using var target = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(target, new System.Text.UTF8Encoding(false));
        writer.Write(xml);
    }

    private static void AssertOffice2021Valid(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
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
