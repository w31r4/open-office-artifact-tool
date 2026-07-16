using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using Xunit;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace OpenChestnut.Codec.Tests;

public sealed class XlsxSparklineCodecTests
{
    [Fact]
    public void AuthorsImportsAndSourceEditsCanonicalSparklineGroups()
    {
        var authored = Invoke(ExportRequest());
        Assert.True(authored.Ok, Diagnostics(authored));
        AssertOffice2021Valid(authored.File.ToByteArray());
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            var worksheet = Assert.Single(document.WorkbookPart!.WorksheetParts).Worksheet!;
            var extension = Assert.Single(worksheet.GetFirstChild<WorksheetExtensionList>()!.Elements<WorksheetExtension>());
            Assert.Equal(XlsxSparklineCodec.ExtensionUri, extension.Uri!.Value, ignoreCase: true);
            var native = Assert.Single(Assert.Single(extension.Elements<X14.SparklineGroups>()).Elements<X14.SparklineGroup>());
            Assert.Equal(X14.SparklineTypeValues.Line, native.Type!.Value);
            Assert.Equal(2.25, native.LineWeight!.Value);
            Assert.True(native.Markers!.Value);
            Assert.True(native.High!.Value);
            Assert.True(native.DisplayHidden!.Value);
            Assert.Equal("FF112233", native.SeriesColor!.Rgb!.Value);
            Assert.Equal("Data!A3:D3", native.Formula!.Text);
            Assert.Collection(native.Sparklines!.Elements<X14.Sparkline>(),
                item =>
                {
                    Assert.Equal("Data!A1:D1", item.Formula!.Text);
                    Assert.Equal("E1", item.ReferenceSequence!.Text);
                },
                item =>
                {
                    Assert.Equal("Data!A2:D2", item.Formula!.Text);
                    Assert.Equal("E2", item.ReferenceSequence!.Text);
                });
        }

        var imported = Import(authored.File.ToByteArray());
        Assert.True(imported.Ok, Diagnostics(imported));
        var group = Assert.Single(Assert.Single(imported.Artifact.Workbook.Worksheets).SparklineGroups);
        Assert.Equal("E1:E2", group.TargetRange);
        Assert.Equal("Data!A1:D2", group.SourceDataRange);
        Assert.Equal("Data!A3:D3", group.DateAxisRange);
        Assert.Equal(SpreadsheetSparklineType.Line, group.Type);
        Assert.Equal("112233", group.SeriesColor.Rgb);
        Assert.True(group.Source.Editable);
        Assert.NotEmpty(group.Source.SemanticSha256);

        group.SeriesColor.Rgb = "445566";
        group.LineWeight = 3;
        group.Axis.ManualMax = 20;
        group.Axis.MaxMode = SpreadsheetSparklineAxisMode.Custom;
        var edited = Export(imported.Artifact);
        Assert.True(edited.Ok, Diagnostics(edited));
        AssertOffice2021Valid(edited.File.ToByteArray());
        var reimported = Import(edited.File.ToByteArray());
        Assert.True(reimported.Ok, Diagnostics(reimported));
        var result = Assert.Single(Assert.Single(reimported.Artifact.Workbook.Worksheets).SparklineGroups);
        Assert.Equal("445566", result.SeriesColor.Rgb);
        Assert.Equal(3, result.LineWeight);
        Assert.Equal(20, result.Axis.ManualMax);
        Assert.Equal(SpreadsheetSparklineAxisMode.Custom, result.Axis.MaxMode);
    }

    [Fact]
    public void RejectsNonReversibleSparklineRangeProfiles()
    {
        var request = ExportRequest();
        request.Artifact.Workbook.Worksheets[0].SparklineGroups[0].TargetRange = "E1:F2";
        var response = Invoke(request);
        Assert.False(response.Ok);
        Assert.Equal("invalid_spreadsheet_sparkline", Assert.Single(response.Diagnostics).Code);
        Assert.Contains("one-dimensional", response.Diagnostics[0].Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AuthorsAndImportsHorizontalTargetColumnMappings()
    {
        var request = ExportRequest();
        var group = request.Artifact.Workbook.Worksheets[0].SparklineGroups[0];
        group.TargetRange = "E1:F1";
        group.SourceDataRange = "Data!A1:B3";
        group.DateAxisRange = "";
        group.Axis.MinMode = SpreadsheetSparklineAxisMode.Group;
        group.Axis.MaxMode = SpreadsheetSparklineAxisMode.Group;
        group.Axis.ClearManualMin();
        group.Axis.ClearManualMax();

        var authored = Invoke(request);
        Assert.True(authored.Ok, Diagnostics(authored));
        AssertOffice2021Valid(authored.File.ToByteArray());
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            var native = Assert.Single(Assert.Single(document.WorkbookPart!.WorksheetParts).Worksheet!
                .GetFirstChild<WorksheetExtensionList>()!.Descendants<X14.SparklineGroup>());
            Assert.Collection(native.Sparklines!.Elements<X14.Sparkline>(),
                item =>
                {
                    Assert.Equal("Data!A1:A3", item.Formula!.Text);
                    Assert.Equal("E1", item.ReferenceSequence!.Text);
                },
                item =>
                {
                    Assert.Equal("Data!B1:B3", item.Formula!.Text);
                    Assert.Equal("F1", item.ReferenceSequence!.Text);
                });
        }

        var imported = Import(authored.File.ToByteArray());
        Assert.True(imported.Ok, Diagnostics(imported));
        var result = Assert.Single(Assert.Single(imported.Artifact.Workbook.Worksheets).SparklineGroups);
        Assert.Equal("E1:F1", result.TargetRange);
        Assert.Equal("Data!A1:B3", result.SourceDataRange);
        Assert.Equal(SpreadsheetSparklineAxisMode.Group, result.Axis.MinMode);
        Assert.Equal(SpreadsheetSparklineAxisMode.Group, result.Axis.MaxMode);
    }

    [Fact]
    public void NonContiguousNativeGroupsStayOpaqueAndRejectSemanticReplacement()
    {
        var authored = Invoke(ExportRequest());
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = MakeSourceGroupNonContiguous(authored.File.ToByteArray());
        AssertOffice2021Valid(source);
        var original = SparklineXml(source);

        var imported = Import(source);
        Assert.True(imported.Ok, Diagnostics(imported));
        var sheet = Assert.Single(imported.Artifact.Workbook.Worksheets);
        Assert.Empty(sheet.SparklineGroups);

        var preserved = Export(imported.Artifact);
        Assert.True(preserved.Ok, Diagnostics(preserved));
        Assert.Equal(original, SparklineXml(preserved.File.ToByteArray()));

        sheet.SparklineGroups.Add(CanonicalGroup());
        var rejected = Export(imported.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_sparkline_topology", Assert.Single(rejected.Diagnostics).Code);
    }

    private static CodecRequest ExportRequest()
    {
        var worksheet = new WorksheetArtifact
        {
            Id = "worksheet/1",
            Name = "Data",
            Visibility = SpreadsheetWorksheetVisibility.Visible,
            ShowGridLines = true,
        };
        for (var row = 0; row < 3; row++)
        for (var column = 0; column < 4; column++)
            worksheet.Cells.Add(new CellArtifact { Row = checked((uint)row), Column = checked((uint)column), NumberValue = row * 10 + column + 1 });
        worksheet.SparklineGroups.Add(CanonicalGroup());
        var workbook = new WorkbookArtifact { Id = "workbook/1", DateSystem = WorkbookDateSystem._1900 };
        workbook.Worksheets.Add(worksheet);
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

    private static SpreadsheetSparklineGroupArtifact CanonicalGroup() => new()
    {
        Id = "sparkline/1",
        Type = SpreadsheetSparklineType.Line,
        TargetRange = "E1:E2",
        SourceDataRange = "Data!A1:D2",
        DateAxisRange = "Data!A3:D3",
        SeriesColor = new SpreadsheetColor { Rgb = "112233" },
        NegativeColor = new SpreadsheetColor { Rgb = "AA0000" },
        MarkersColor = new SpreadsheetColor { Theme = 4, Tint = 0.2 },
        LineWeight = 2.25,
        DisplayHidden = true,
        DisplayEmptyCellsAs = SpreadsheetSparklineEmptyCells.Gap,
        Markers = new SpreadsheetSparklineMarkersArtifact { Show = true, High = true, Negative = true },
        Axis = new SpreadsheetSparklineAxisArtifact
        {
            ManualMin = -5,
            ManualMax = 10,
            MinMode = SpreadsheetSparklineAxisMode.Custom,
            MaxMode = SpreadsheetSparklineAxisMode.Custom,
            ShowAxis = true,
            RightToLeft = false,
        },
    };

    private static CodecResponse Import(byte[] file) => Invoke(new CodecRequest
    {
        ProtocolVersion = CodecProtocol.ProtocolVersion,
        Operation = CodecOperation.ImportXlsx,
        Family = ArtifactFamily.Workbook,
        File = ByteString.CopyFrom(file),
    });

    private static CodecResponse Export(ArtifactEnvelope artifact) => Invoke(new CodecRequest
    {
        ProtocolVersion = CodecProtocol.ProtocolVersion,
        Operation = CodecOperation.ExportXlsx,
        Family = ArtifactFamily.Workbook,
        Artifact = artifact,
    });

    private static CodecResponse Invoke(CodecRequest request) => CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
    private static string Diagnostics(CodecResponse response) => string.Join("\n", response.Diagnostics.Select(item => $"{item.Code}: {item.Message}"));

    private static void AssertOffice2021Valid(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
    }

    private static byte[] MakeSourceGroupNonContiguous(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var native = Assert.Single(Assert.Single(Assert.Single(document.WorkbookPart!.WorksheetParts).Worksheet!
                .GetFirstChild<WorksheetExtensionList>()!.Descendants<X14.SparklineGroups>()).Elements<X14.SparklineGroup>());
            native.Sparklines!.Elements<X14.Sparkline>().Last().Formula!.Text = "Data!A4:D4";
            native.Ancestors<Worksheet>().Single().Save();
        }
        return stream.ToArray();
    }

    private static string SparklineXml(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        return Assert.Single(Assert.Single(document.WorkbookPart!.WorksheetParts).Worksheet!
            .GetFirstChild<WorksheetExtensionList>()!.Descendants<X14.SparklineGroups>()).OuterXml;
    }
}
