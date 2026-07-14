using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using System.IO.Compression;
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
    }

    [Fact]
    public void ImportRejectsDataTableFormulaAndFormulaEditRejectsUnmodeledAttributes()
    {
        var first = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(FormulaExportRequest().ToByteArray()));
        var dataTable = Import(SetFormulaType(first.File.ToByteArray(), "C1", CellFormulaValues.DataTable));
        Assert.False(dataTable.Ok);
        Assert.Equal("unsupported_cell_formula", Assert.Single(dataTable.Diagnostics).Code);

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
