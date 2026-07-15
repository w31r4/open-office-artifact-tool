using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml.Linq;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

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
    public void ProtocolAuthorsImportsAndSourcePreservesWorkbookCalculationPolicy()
    {
        var request = ExportRequest();
        request.Artifact.Workbook.Calculation = new SpreadsheetCalculationArtifact
        {
            Mode = SpreadsheetCalculationMode.AutomaticExceptTables,
            CalculateOnSave = false,
            FullCalculationOnLoad = true,
            ForceFullCalculation = true,
            IterationEnabled = true,
            MaxIterations = 100,
            MaxChange = 0.001,
            FullPrecision = false,
        };
        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.True(authored.Ok, string.Join("\n", authored.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(authored.File.ToByteArray());
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            var calculation = document.WorkbookPart!.Workbook!.CalculationProperties!;
            Assert.Equal(CalculateModeValues.AutoNoTable, calculation.CalculationMode?.Value);
            Assert.False(calculation.CalculationOnSave!.Value);
            Assert.True(calculation.FullCalculationOnLoad!.Value);
            Assert.True(calculation.ForceFullCalculation!.Value);
            Assert.True(calculation.Iterate!.Value);
            Assert.Equal(100U, calculation.IterateCount?.Value);
            Assert.Equal(0.001, calculation.IterateDelta?.Value);
            Assert.False(calculation.FullPrecision!.Value);
        }

        var source = SetCalculationProfile(authored.File.ToByteArray(), calculation => calculation.CalculationId = 191029U);
        var imported = Import(source);
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var policy = imported.Artifact.Workbook.Calculation;
        Assert.NotNull(policy);
        Assert.Equal(SpreadsheetCalculationMode.AutomaticExceptTables, policy.Mode);
        Assert.True(policy.HasCalculateOnSave);
        Assert.False(policy.CalculateOnSave);
        Assert.Equal(100U, policy.MaxIterations);
        Assert.Equal(0.001, policy.MaxChange);
        Assert.True(policy.Source.Editable);
        Assert.Equal(64, policy.Source.WorkbookXmlSha256.Length);

        policy.Mode = SpreadsheetCalculationMode.Manual;
        policy.ForceFullCalculation = false;
        policy.MaxIterations = 250;
        policy.MaxChange = 0.0001;
        var preserved = Export(imported.Artifact);
        Assert.True(preserved.Ok, string.Join("\n", preserved.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        using var preservedStream = new MemoryStream(preserved.File.ToByteArray());
        using var preservedDocument = SpreadsheetDocument.Open(preservedStream, false);
        var edited = preservedDocument.WorkbookPart!.Workbook!.CalculationProperties!;
        Assert.Equal(191029U, edited.CalculationId?.Value);
        Assert.Equal(CalculateModeValues.Manual, edited.CalculationMode?.Value);
        Assert.False(edited.ForceFullCalculation!.Value);
        Assert.Equal(250U, edited.IterateCount?.Value);
        Assert.Equal(0.0001, edited.IterateDelta?.Value);
    }

    [Fact]
    public void WorkbookCalculationPolicyRejectsInvalidTopologyAndPreservesOpaqueProfiles()
    {
        var invalid = ExportRequest();
        invalid.Artifact.Workbook.Calculation = new SpreadsheetCalculationArtifact { Mode = SpreadsheetCalculationMode.Automatic, MaxIterations = 0 };
        var response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(invalid.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_workbook_calculation", Assert.Single(response.Diagnostics).Code);

        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(ExportRequest().ToByteArray()));
        var absent = Import(authored.File.ToByteArray());
        absent.Artifact.Workbook.Calculation = new SpreadsheetCalculationArtifact { Mode = SpreadsheetCalculationMode.Automatic };
        response = Export(absent.Artifact);
        Assert.False(response.Ok);
        Assert.Contains("cannot add", Assert.Single(response.Diagnostics).Message, StringComparison.OrdinalIgnoreCase);

        var boundedRequest = ExportRequest();
        boundedRequest.Artifact.Workbook.Calculation = new SpreadsheetCalculationArtifact { Mode = SpreadsheetCalculationMode.Automatic };
        var boundedSource = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(boundedRequest.ToByteArray()));
        var bounded = Import(boundedSource.File.ToByteArray());
        bounded.Artifact.Workbook.Calculation.Source.CalculationXmlSha256 = new string('0', 64);
        response = Export(bounded.Artifact);
        Assert.False(response.Ok);
        Assert.Contains("source binding", Assert.Single(response.Diagnostics).Message, StringComparison.OrdinalIgnoreCase);

        bounded = Import(boundedSource.File.ToByteArray());
        bounded.Artifact.Workbook.Calculation = null;
        response = Export(bounded.Artifact);
        Assert.False(response.Ok);
        Assert.Contains("cannot remove", Assert.Single(response.Diagnostics).Message, StringComparison.OrdinalIgnoreCase);

        var opaqueSource = SetCalculationProfile(authored.File.ToByteArray(), calculation =>
        {
            calculation.CalculationMode = CalculateModeValues.Auto;
            calculation.ReferenceMode = ReferenceModeValues.R1C1;
            calculation.CalculationId = 191029U;
        });
        var opaqueXml = ReadCalculationXml(opaqueSource);
        var opaque = Import(opaqueSource);
        Assert.True(opaque.Ok, string.Join("\n", opaque.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Null(opaque.Artifact.Workbook.Calculation);
        var roundtrip = Export(opaque.Artifact);
        Assert.True(roundtrip.Ok, string.Join("\n", roundtrip.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal(opaqueXml, ReadCalculationXml(roundtrip.File.ToByteArray()));

        opaque.Artifact.Workbook.Calculation = new SpreadsheetCalculationArtifact { Mode = SpreadsheetCalculationMode.Manual };
        response = Export(opaque.Artifact);
        Assert.False(response.Ok);
        Assert.Contains("opaque workbook calculation profile", Assert.Single(response.Diagnostics).Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ProtocolAuthorsImportsAndSourcePreservesWorksheetVisibility()
    {
        var request = ExportRequest();
        request.Artifact.Workbook.Worksheets[0].Visibility = SpreadsheetWorksheetVisibility.Visible;
        request.Artifact.Workbook.Worksheets.Add(new WorksheetArtifact
        {
            Id = "worksheet/hidden",
            Name = "Hidden Data",
            Visibility = SpreadsheetWorksheetVisibility.Hidden,
            ShowGridLines = true,
            Cells = { new CellArtifact { Row = 0, Column = 0, StringValue = "hidden" } },
        });
        request.Artifact.Workbook.Worksheets.Add(new WorksheetArtifact
        {
            Id = "worksheet/very-hidden",
            Name = "Internal State",
            Visibility = SpreadsheetWorksheetVisibility.VeryHidden,
            ShowGridLines = true,
            Cells = { new CellArtifact { Row = 0, Column = 0, StringValue = "internal" } },
        });

        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.True(authored.Ok, string.Join("\n", authored.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(authored.File.ToByteArray());
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            var sheets = document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().ToArray();
            Assert.Null(sheets[0].State);
            Assert.Equal(SheetStateValues.Hidden, sheets[1].State?.Value);
            Assert.Equal(SheetStateValues.VeryHidden, sheets[2].State?.Value);
            Assert.Equal(0U, document.WorkbookPart.Workbook.BookViews!.Elements<WorkbookView>().Single().ActiveTab?.Value);
        }

        var imported = Import(authored.File.ToByteArray());
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Collection(imported.Artifact.Workbook.Worksheets,
            sheet =>
            {
                Assert.Equal(SpreadsheetWorksheetVisibility.Visible, sheet.Visibility);
                Assert.True(sheet.Source.Editable);
                Assert.Equal(0U, sheet.Source.Ordinal);
            },
            sheet =>
            {
                Assert.Equal(SpreadsheetWorksheetVisibility.Hidden, sheet.Visibility);
                Assert.Equal(1U, sheet.Source.Ordinal);
                Assert.Equal(64, sheet.Source.SheetElementSha256.Length);
            },
            sheet => Assert.Equal(SpreadsheetWorksheetVisibility.VeryHidden, sheet.Visibility));

        imported.Artifact.Workbook.Worksheets[1].Name = "Hidden Archive";
        imported.Artifact.Workbook.Worksheets[1].Visibility = SpreadsheetWorksheetVisibility.VeryHidden;
        imported.Artifact.Workbook.Worksheets[2].Visibility = SpreadsheetWorksheetVisibility.Hidden;
        var edited = Export(imported.Artifact);
        Assert.True(edited.Ok, string.Join("\n", edited.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(edited.File.ToByteArray());
        var reimported = Import(edited.File.ToByteArray());
        Assert.Equal("Hidden Archive", reimported.Artifact.Workbook.Worksheets[1].Name);
        Assert.Equal(SpreadsheetWorksheetVisibility.VeryHidden, reimported.Artifact.Workbook.Worksheets[1].Visibility);
        Assert.Equal(SpreadsheetWorksheetVisibility.Hidden, reimported.Artifact.Workbook.Worksheets[2].Visibility);
    }

    [Fact]
    public void WorksheetVisibilityRejectsAllHiddenActiveOpaqueAndTamperedProfiles()
    {
        var allHidden = ExportRequest();
        allHidden.Artifact.Workbook.Worksheets[0].Visibility = SpreadsheetWorksheetVisibility.Hidden;
        var response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(allHidden.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_worksheet_metadata", Assert.Single(response.Diagnostics).Code);
        Assert.Contains("at least one visible", response.Diagnostics[0].Message, StringComparison.OrdinalIgnoreCase);

        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(ExportRequest().ToByteArray()));
        var imported = Import(authored.File.ToByteArray());
        imported.Artifact.Workbook.Worksheets[0].Visibility = SpreadsheetWorksheetVisibility.Hidden;
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Contains("at least one visible", Assert.Single(response.Diagnostics).Message, StringComparison.OrdinalIgnoreCase);

        var twoSheets = ExportRequest();
        twoSheets.Artifact.Workbook.Worksheets.Add(new WorksheetArtifact
        {
            Id = "worksheet/second",
            Name = "Second",
            Visibility = SpreadsheetWorksheetVisibility.Visible,
            Cells = { new CellArtifact { Row = 0, Column = 0, StringValue = "second" } },
        });
        var twoSheetSource = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(twoSheets.ToByteArray()));
        imported = Import(twoSheetSource.File.ToByteArray());
        imported.Artifact.Workbook.Worksheets[0].Visibility = SpreadsheetWorksheetVisibility.Hidden;
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Contains("active worksheet", Assert.Single(response.Diagnostics).Message, StringComparison.OrdinalIgnoreCase);

        imported = Import(twoSheetSource.File.ToByteArray());
        imported.Artifact.Workbook.Worksheets[1].Source.SheetElementSha256 = new string('0', 64);
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Contains("source binding", Assert.Single(response.Diagnostics).Message, StringComparison.OrdinalIgnoreCase);

        var opaqueSource = SetRawSheetState(twoSheetSource.File.ToByteArray(), "Second", "futureHidden");
        imported = Import(opaqueSource);
        Assert.Equal(SpreadsheetWorksheetVisibility.Unspecified, imported.Artifact.Workbook.Worksheets[1].Visibility);
        Assert.False(imported.Artifact.Workbook.Worksheets[1].Source.Editable);
        imported.Artifact.Workbook.Worksheets[1].Visibility = SpreadsheetWorksheetVisibility.Hidden;
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Contains("opaque worksheet metadata", Assert.Single(response.Diagnostics).Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ProtocolAuthorsImportsAndSourcePreservesActiveWorksheetSelection()
    {
        var request = ExportRequest();
        request.Artifact.Workbook.Worksheets.Add(new WorksheetArtifact
        {
            Id = "worksheet/detail",
            Name = "Detail",
            Visibility = SpreadsheetWorksheetVisibility.Visible,
            ShowGridLines = true,
            Cells = { new CellArtifact { Row = 0, Column = 0, StringValue = "detail" } },
        });
        request.Artifact.Workbook.View = new SpreadsheetWorkbookViewArtifact { ActiveWorksheetId = "worksheet/detail" };

        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.True(authored.Ok, string.Join("\n", authored.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(authored.File.ToByteArray());
        Assert.Equal(1U, ReadWorkbookViews(authored.File.ToByteArray()).Single().ActiveTab?.Value);

        var profiled = MutateWorkbookViews(authored.File.ToByteArray(), views =>
        {
            var view = views.Elements<WorkbookView>().Single();
            view.FirstSheet = 1U;
            view.XWindow = 120;
            view.YWindow = 240;
        });
        var imported = Import(profiled);
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal("worksheet/2", imported.Artifact.Workbook.View.ActiveWorksheetId);
        Assert.True(imported.Artifact.Workbook.View.Source.Editable);
        Assert.Equal(64, imported.Artifact.Workbook.View.Source.ViewXmlSha256.Length);

        imported.Artifact.Workbook.View.ActiveWorksheetId = "worksheet/1";
        imported.Artifact.Workbook.Worksheets[1].Visibility = SpreadsheetWorksheetVisibility.Hidden;
        var edited = Export(imported.Artifact);
        Assert.True(edited.Ok, string.Join("\n", edited.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(edited.File.ToByteArray());
        var editedView = ReadWorkbookViews(edited.File.ToByteArray()).Single();
        Assert.Equal(0U, editedView.ActiveTab?.Value);
        Assert.Equal(1U, editedView.FirstSheet?.Value);
        Assert.Equal(120, editedView.XWindow?.Value);
        Assert.Equal(240, editedView.YWindow?.Value);
        var reimported = Import(edited.File.ToByteArray());
        Assert.Equal("worksheet/1", reimported.Artifact.Workbook.View.ActiveWorksheetId);
        Assert.Equal(SpreadsheetWorksheetVisibility.Hidden, reimported.Artifact.Workbook.Worksheets[1].Visibility);

        imported = Import(profiled);
        imported.Artifact.Workbook.View.Source.ViewXmlSha256 = new string('0', 64);
        var rejected = Export(imported.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_workbook_view", Assert.Single(rejected.Diagnostics).Code);
        Assert.Contains("source binding", rejected.Diagnostics[0].Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void WorkbookViewRejectsOpaqueMultiWindowEditsAndPreservesTheGraph()
    {
        var request = ExportRequest();
        request.Artifact.Workbook.Worksheets.Add(new WorksheetArtifact
        {
            Id = "worksheet/detail",
            Name = "Detail",
            Visibility = SpreadsheetWorksheetVisibility.Visible,
            ShowGridLines = true,
        });
        request.Artifact.Workbook.View = new SpreadsheetWorkbookViewArtifact { ActiveWorksheetId = "worksheet/detail" };
        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        var multiWindow = MutateWorkbookViews(authored.File.ToByteArray(), views => views.Append(new WorkbookView { ActiveTab = 0U, XWindow = 777 }));

        var imported = Import(multiWindow);
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Null(imported.Artifact.Workbook.View);
        var preserved = Export(imported.Artifact);
        Assert.True(preserved.Ok, string.Join("\n", preserved.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Collection(ReadWorkbookViews(preserved.File.ToByteArray()),
            view => Assert.Equal(1U, view.ActiveTab?.Value),
            view =>
            {
                Assert.Equal(0U, view.ActiveTab?.Value);
                Assert.Equal(777, view.XWindow?.Value);
            });

        imported = Import(multiWindow);
        imported.Artifact.Workbook.View = new SpreadsheetWorkbookViewArtifact { ActiveWorksheetId = "worksheet/1" };
        var rejected = Export(imported.Artifact);
        Assert.False(rejected.Ok);
        Assert.Contains("multi-window", Assert.Single(rejected.Diagnostics).Message, StringComparison.OrdinalIgnoreCase);

        imported = Import(multiWindow);
        imported.Artifact.Workbook.Worksheets[1].Visibility = SpreadsheetWorksheetVisibility.Hidden;
        rejected = Export(imported.Artifact);
        Assert.False(rejected.Ok);
        Assert.Contains("active worksheet", Assert.Single(rejected.Diagnostics).Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void WorkbookViewAuthorsImportsAndSourcePreservesMultipleWindows()
    {
        var request = ExportRequest();
        request.Artifact.Workbook.Worksheets.Add(new WorksheetArtifact
        {
            Id = "worksheet/detail",
            Name = "Detail",
            Visibility = SpreadsheetWorksheetVisibility.Visible,
            ShowGridLines = true,
        });
        request.Artifact.Workbook.Worksheets.Add(new WorksheetArtifact
        {
            Id = "worksheet/review",
            Name = "Review",
            Visibility = SpreadsheetWorksheetVisibility.Visible,
            ShowGridLines = true,
        });
        request.Artifact.Workbook.View = new SpreadsheetWorkbookViewArtifact { ActiveWorksheetId = "worksheet/detail" };
        request.Artifact.Workbook.View.SelectedWorksheetIds.Add(["worksheet/summary", "worksheet/detail"]);
        var additional = new SpreadsheetWorkbookViewArtifact { ActiveWorksheetId = "worksheet/review" };
        additional.SelectedWorksheetIds.Add(["worksheet/detail", "worksheet/review"]);
        request.Artifact.Workbook.AdditionalViews.Add(additional);

        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.True(authored.Ok, string.Join("\n", authored.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(authored.File.ToByteArray());
        Assert.Collection(ReadWorkbookViews(authored.File.ToByteArray()),
            view => Assert.Equal(1U, view.ActiveTab?.Value),
            view => Assert.Equal(2U, view.ActiveTab?.Value));
        var authoredSheetViews = ReadWorksheetViewMatrix(authored.File.ToByteArray());
        Assert.Equal([true, false], authoredSheetViews[0].Select(view => view.TabSelected?.Value ?? false));
        Assert.Equal([true, true], authoredSheetViews[1].Select(view => view.TabSelected?.Value ?? false));
        Assert.Equal([false, true], authoredSheetViews[2].Select(view => view.TabSelected?.Value ?? false));

        var profiled = MutateWorkbookViews(authored.File.ToByteArray(), views =>
        {
            views.Elements<WorkbookView>().ElementAt(0).XWindow = 111;
            views.Elements<WorkbookView>().ElementAt(1).YWindow = 222;
        });
        profiled = MutateWorksheetViews(profiled, 1, views => views.Elements<SheetView>().ElementAt(1).ZoomScale = 85U);
        profiled = MutateWorksheetViews(profiled, 0, views => views.Elements<SheetView>().ElementAt(1).TabSelected = false);
        var imported = Import(profiled);
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal("worksheet/2", imported.Artifact.Workbook.View.ActiveWorksheetId);
        Assert.Equal(["worksheet/1", "worksheet/2"], imported.Artifact.Workbook.View.SelectedWorksheetIds);
        var importedAdditional = Assert.Single(imported.Artifact.Workbook.AdditionalViews);
        Assert.Equal("worksheet/3", importedAdditional.ActiveWorksheetId);
        Assert.Equal(["worksheet/2", "worksheet/3"], importedAdditional.SelectedWorksheetIds);
        Assert.Equal(0U, imported.Artifact.Workbook.View.Source.Ordinal);
        Assert.Equal(1U, importedAdditional.Source.Ordinal);
        Assert.All(importedAdditional.Source.WorksheetViews, binding => Assert.Equal(1U, binding.WorkbookViewId));

        importedAdditional.ActiveWorksheetId = "worksheet/2";
        importedAdditional.SelectedWorksheetIds.Clear();
        importedAdditional.SelectedWorksheetIds.Add("worksheet/2");
        var edited = Export(imported.Artifact);
        Assert.True(edited.Ok, string.Join("\n", edited.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(edited.File.ToByteArray());
        Assert.Collection(ReadWorkbookViews(edited.File.ToByteArray()),
            view =>
            {
                Assert.Equal(1U, view.ActiveTab?.Value);
                Assert.Equal(111, view.XWindow?.Value);
            },
            view =>
            {
                Assert.Equal(1U, view.ActiveTab?.Value);
                Assert.Equal(222, view.YWindow?.Value);
            });
        var editedSheetViews = ReadWorksheetViewMatrix(edited.File.ToByteArray());
        Assert.Equal([true, false], editedSheetViews[0].Select(view => view.TabSelected?.Value ?? false));
        Assert.NotNull(editedSheetViews[0][1].TabSelected);
        Assert.Equal([true, true], editedSheetViews[1].Select(view => view.TabSelected?.Value ?? false));
        Assert.Equal([false, false], editedSheetViews[2].Select(view => view.TabSelected?.Value ?? false));
        Assert.Equal(85U, editedSheetViews[1][1].ZoomScale?.Value);

        imported = Import(profiled);
        imported.Artifact.Workbook.AdditionalViews.Clear();
        var rejected = Export(imported.Artifact);
        Assert.False(rejected.Ok);
        Assert.Contains("window count or order", Assert.Single(rejected.Diagnostics).Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void WorkbookViewAuthorsImportsAndSourcePreservesGroupedTabSelection()
    {
        var request = ExportRequest();
        request.Artifact.Workbook.Worksheets.Add(new WorksheetArtifact
        {
            Id = "worksheet/detail",
            Name = "Detail",
            Visibility = SpreadsheetWorksheetVisibility.Visible,
            ShowGridLines = true,
        });
        request.Artifact.Workbook.View = new SpreadsheetWorkbookViewArtifact { ActiveWorksheetId = "worksheet/detail" };
        request.Artifact.Workbook.View.SelectedWorksheetIds.Add(["worksheet/summary", "worksheet/detail"]);

        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.True(authored.Ok, string.Join("\n", authored.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.All(ReadWorksheetViews(authored.File.ToByteArray()), view => Assert.True(view!.TabSelected?.Value));

        var profiled = MutateWorksheetViews(authored.File.ToByteArray(), 0, views =>
        {
            var view = Assert.Single(views.Elements<SheetView>());
            view.TabSelected = false;
            view.TopLeftCell = "C4";
            view.ZoomScale = 125U;
        });
        var imported = Import(profiled);
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal("worksheet/2", imported.Artifact.Workbook.View.ActiveWorksheetId);
        Assert.Equal(["worksheet/2"], imported.Artifact.Workbook.View.SelectedWorksheetIds);
        Assert.True(imported.Artifact.Workbook.View.Source.Editable);
        Assert.Equal(2, imported.Artifact.Workbook.View.Source.WorksheetViews.Count);
        Assert.True(imported.Artifact.Workbook.View.Source.WorksheetViews[0].HasTabSelected);
        Assert.False(imported.Artifact.Workbook.View.Source.WorksheetViews[0].TabSelected);
        Assert.All(imported.Artifact.Workbook.View.Source.WorksheetViews, binding =>
        {
            Assert.Equal(64, binding.WorksheetXmlSha256.Length);
            Assert.Equal(64, binding.ViewXmlSha256.Length);
            Assert.Equal(0U, binding.WorkbookViewId);
            Assert.True(binding.Editable);
        });

        var preserved = Export(imported.Artifact);
        Assert.True(preserved.Ok, string.Join("\n", preserved.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.False(ReadWorksheetViews(preserved.File.ToByteArray())[0]!.TabSelected?.Value);

        imported.Artifact.Workbook.View.SelectedWorksheetIds.Insert(0, "worksheet/1");
        var edited = Export(imported.Artifact);
        Assert.True(edited.Ok, string.Join("\n", edited.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(edited.File.ToByteArray());
        var editedViews = ReadWorksheetViews(edited.File.ToByteArray());
        Assert.True(editedViews[0]!.TabSelected?.Value);
        Assert.True(editedViews[1]!.TabSelected?.Value);
        Assert.Equal("C4", editedViews[0]!.TopLeftCell?.Value);
        Assert.Equal(125U, editedViews[0]!.ZoomScale?.Value);
        var reimported = Import(edited.File.ToByteArray());
        Assert.Equal(["worksheet/1", "worksheet/2"], reimported.Artifact.Workbook.View.SelectedWorksheetIds);

        imported = Import(profiled);
        imported.Artifact.Workbook.View.Source.WorksheetViews[0].ViewXmlSha256 = new string('0', 64);
        var rejected = Export(imported.Artifact);
        Assert.False(rejected.Ok);
        Assert.Contains("worksheet source binding", Assert.Single(rejected.Diagnostics).Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void WorkbookViewKeepsIncompleteWorksheetSelectionOpaqueAndFailClosed()
    {
        var request = ExportRequest();
        request.Artifact.Workbook.Worksheets.Add(new WorksheetArtifact
        {
            Id = "worksheet/detail",
            Name = "Detail",
            Visibility = SpreadsheetWorksheetVisibility.Visible,
            ShowGridLines = true,
        });
        request.Artifact.Workbook.View = new SpreadsheetWorkbookViewArtifact { ActiveWorksheetId = "worksheet/detail" };
        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        var incomplete = MutateWorksheetViews(authored.File.ToByteArray(), 0, views => views.Remove());

        var imported = Import(incomplete);
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.NotNull(imported.Artifact.Workbook.View);
        Assert.False(imported.Artifact.Workbook.View.Source.Editable);
        Assert.Empty(imported.Artifact.Workbook.View.SelectedWorksheetIds);
        Assert.Empty(imported.Artifact.Workbook.View.Source.WorksheetViews);

        var preserved = Export(imported.Artifact);
        Assert.True(preserved.Ok, string.Join("\n", preserved.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(preserved.File.ToByteArray());
        Assert.Null(ReadWorksheetViews(preserved.File.ToByteArray())[0]);
        Assert.True(ReadWorksheetViews(preserved.File.ToByteArray())[1]!.TabSelected?.Value);

        imported = Import(incomplete);
        imported.Artifact.Workbook.View.SelectedWorksheetIds.Add(["worksheet/1", "worksheet/2"]);
        var rejected = Export(imported.Artifact);
        Assert.False(rejected.Ok);
        Assert.Contains("opaque worksheet-selection", Assert.Single(rejected.Diagnostics).Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ProtocolAuthorsImportsAndPreservesWorkbookAndWorksheetDefinedNames()
    {
        var request = DefinedNameExportRequest();
        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(exported.File.ToByteArray());

        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            var names = document.WorkbookPart!.Workbook!.DefinedNames!.Elements<DefinedName>().ToArray();
            Assert.Collection(names,
                item =>
                {
                    Assert.Equal("RevenueData", item.Name?.Value);
                    Assert.Equal("Summary!$B$1:$B$2", item.Text);
                    Assert.Equal("Revenue body", item.Comment?.Value);
                    Assert.Null(item.LocalSheetId);
                    Assert.NotNull(item.Hidden);
                    Assert.False(item.Hidden!.Value);
                },
                item =>
                {
                    Assert.Equal("RevenueData", item.Name?.Value);
                    Assert.Equal(0U, item.LocalSheetId?.Value);
                    Assert.Null(item.Comment);
                    Assert.True(item.Hidden?.Value);
                });
        }

        var imported = Import(exported.File.ToByteArray());
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Collection(imported.Artifact.Workbook.DefinedNames,
            item =>
            {
                Assert.Equal("RevenueData", item.Name);
                Assert.Equal("Revenue body", item.Comment);
                Assert.True(item.HasHidden);
                Assert.False(item.Hidden);
                Assert.False(item.HasScopeSheetName);
                Assert.True(item.Source.Editable);
                Assert.Equal(0U, item.Source.Ordinal);
            },
            item =>
            {
                Assert.Equal("Summary", item.ScopeSheetName);
                Assert.True(item.Hidden);
                Assert.Equal(1U, item.Source.Ordinal);
            });
    }

    [Fact]
    public void SourcePreservingDefinedNameEditKeepsOpaqueNamesAndRejectsTopologyChanges()
    {
        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(DefinedNameExportRequest().ToByteArray()));
        Assert.True(authored.Ok, string.Join("\n", authored.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var source = AddOpaqueDefinedName(authored.File.ToByteArray());
        var opaqueXml = ReadDefinedNameXml(source, "OpaqueConstant");
        var imported = Import(source);
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal(2, imported.Artifact.Workbook.DefinedNames.Count);

        var edited = imported.Artifact.Workbook.DefinedNames[0];
        edited.Name = "RevenueRange";
        edited.RefersTo = "Summary!$B$1";
        edited.Comment = "Updated range";
        edited.Hidden = true;
        var preserved = Export(imported.Artifact);
        Assert.True(preserved.Ok, string.Join("\n", preserved.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(preserved.File.ToByteArray());
        Assert.Equal(opaqueXml, ReadDefinedNameXml(preserved.File.ToByteArray(), "OpaqueConstant"));
        var reimported = Import(preserved.File.ToByteArray());
        Assert.True(reimported.Ok, string.Join("\n", reimported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal("RevenueRange", reimported.Artifact.Workbook.DefinedNames[0].Name);
        Assert.Equal("Summary!$B$1", reimported.Artifact.Workbook.DefinedNames[0].RefersTo);
        Assert.True(reimported.Artifact.Workbook.DefinedNames[0].Hidden);

        imported = Import(source);
        imported.Artifact.Workbook.DefinedNames.RemoveAt(0);
        var rejected = Export(imported.Artifact);
        Assert.False(rejected.Ok);
        Assert.Contains("add or remove recognized defined names", Assert.Single(rejected.Diagnostics).Message);

        imported = Import(source);
        imported.Artifact.Workbook.DefinedNames[0].Name = "OpaqueConstant";
        rejected = Export(imported.Artifact);
        Assert.False(rejected.Ok);
        Assert.Contains("collides with an opaque", Assert.Single(rejected.Diagnostics).Message);

        imported = Import(source);
        imported.Artifact.Workbook.DefinedNames[0].Source.DefinedNameXmlSha256 = new string('0', 64);
        rejected = Export(imported.Artifact);
        Assert.False(rejected.Ok);
        Assert.Contains("source binding", Assert.Single(rejected.Diagnostics).Message);
    }

    [Fact]
    public void WorkbookDefinedNamesRejectInvalidNamesReferencesScopesAndDuplicates()
    {
        var request = ExportRequest();
        request.Artifact.Workbook.DefinedNames.Add(new SpreadsheetDefinedNameArtifact { Id = "defined-name/1", Name = "A1", RefersTo = "Summary!A1" });
        var response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_workbook_defined_name", Assert.Single(response.Diagnostics).Code);

        request = ExportRequest();
        request.Artifact.Workbook.DefinedNames.Add(new SpreadsheetDefinedNameArtifact { Id = "defined-name/1", Name = "RangeOne", RefersTo = "SUM(Summary!A1:A2)" });
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_workbook_defined_name", Assert.Single(response.Diagnostics).Code);

        request = ExportRequest();
        request.Artifact.Workbook.DefinedNames.Add(new SpreadsheetDefinedNameArtifact { Id = "defined-name/1", Name = "RangeOne", RefersTo = "Summary!A1", ScopeSheetName = "Missing" });
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_workbook_defined_name", Assert.Single(response.Diagnostics).Code);

        request = ExportRequest();
        request.Artifact.Workbook.DefinedNames.Add(new SpreadsheetDefinedNameArtifact { Id = "defined-name/1", Name = "RangeOne", RefersTo = "Summary!A1" });
        request.Artifact.Workbook.DefinedNames.Add(new SpreadsheetDefinedNameArtifact { Id = "defined-name/2", Name = "rangeone", RefersTo = "Summary!B1" });
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Contains("duplicated", Assert.Single(response.Diagnostics).Message);
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
        var relationshipXml = ReadEntry(source, "xl/tables/_rels/table1.xml.rels");

        var imported = Import(source);
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var table = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Tables);
        var connection = Assert.Single(imported.Artifact.Workbook.Connections);
        Assert.Equal(7U, connection.ConnectionId);
        Assert.Equal("Fixture warehouse", connection.Name);
        Assert.Equal("Read-only warehouse source", connection.Description);
        Assert.Equal(5U, connection.Type);
        Assert.Equal(8U, connection.RefreshedVersion);
        Assert.False(connection.KeepAlive);
        Assert.Equal(30U, connection.IntervalMinutes);
        Assert.True(connection.Background);
        Assert.False(connection.RefreshOnLoad);
        Assert.True(connection.SaveData);
        Assert.Equal("xl/connections.xml", connection.Source.PartPath);
        Assert.True(connection.Source.Editable);
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
        Assert.Equal("stroke", refreshSort.SortMethod);
        Assert.False(refreshSort.HasColumnSort);
        Assert.Collection(refreshSort.Conditions,
            condition =>
            {
                Assert.Equal("B2:B3", condition.Reference);
                Assert.True(condition.Descending);
                Assert.Null(condition.Icon);
                Assert.Equal("North,South", condition.CustomList);
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
        refreshSort.SortMethod = "pinYin";
        refreshSort.ColumnSort = true;
        refreshSort.Conditions[0].Reference = "A3:B3";
        refreshSort.Conditions[0].Descending = false;
        refreshSort.Conditions[0].CustomList = "South,North";
        refreshSort.Conditions[1].Reference = "A2:B2";
        refreshSort.Conditions[1].Icon.IconId = 1;
        connection.Name = "Fixture warehouse curated";
        connection.Description = "Curated without executing the source";
        connection.KeepAlive = true;
        connection.IntervalMinutes = 45;
        connection.Background = false;
        connection.RefreshOnLoad = true;
        connection.SaveData = false;
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
        Assert.Equal(relationshipXml, ReadEntry(output, "xl/tables/_rels/table1.xml.rels"));
        var editedConnectionXml = System.Text.Encoding.UTF8.GetString(ReadEntry(output, "xl/connections.xml"));
        Assert.Contains("name=\"Fixture warehouse curated\"", editedConnectionXml);
        Assert.Contains("description=\"Curated without executing the source\"", editedConnectionXml);
        Assert.Contains("keepAlive=\"1\"", editedConnectionXml);
        Assert.Contains("interval=\"45\"", editedConnectionXml);
        Assert.Contains("background=\"0\"", editedConnectionXml);
        Assert.Contains("refreshOnLoad=\"1\"", editedConnectionXml);
        Assert.Contains("saveData=\"0\"", editedConnectionXml);
        Assert.Contains("Provider=Fixture.Provider;Data Source=fixture.invalid", editedConnectionXml);
        Assert.Contains("SELECT Region, Revenue FROM Sales", editedConnectionXml);
        Assert.Contains("savePassword=\"0\"", editedConnectionXml);
        Assert.Contains("credentials=\"integrated\"", editedConnectionXml);
        Assert.Contains("<fixture:connectionOpaque value=\"kept\"", editedConnectionXml);
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
        Assert.Contains("<x:sortState ref=\"A2:B3\" sortMethod=\"pinYin\" columnSort=\"1\">", queryXml);
        Assert.Contains("<x:sortCondition ref=\"A3:B3\" customList=\"South,North\"", queryXml);
        Assert.Contains("<x:sortCondition ref=\"A2:B2\" sortBy=\"icon\" iconSet=\"3Arrows\" iconId=\"1\"", queryXml);
        Assert.Contains("<fixture:fieldOpaque value=\"kept\"", queryXml);
        Assert.Contains("<fixture:sortOpaque value=\"kept\"", queryXml);
        Assert.Contains("<fixture:opaque value=\"kept\"", queryXml);

        var reimported = Import(output);
        var edited = Assert.Single(reimported.Artifact.Workbook.Worksheets[0].Tables);
        var editedConnection = Assert.Single(reimported.Artifact.Workbook.Connections);
        Assert.Equal("Fixture warehouse curated", editedConnection.Name);
        Assert.Equal("Curated without executing the source", editedConnection.Description);
        Assert.True(editedConnection.KeepAlive);
        Assert.Equal(45U, editedConnection.IntervalMinutes);
        Assert.False(editedConnection.Background);
        Assert.True(editedConnection.RefreshOnLoad);
        Assert.False(editedConnection.SaveData);
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
        Assert.Equal("pinYin", edited.QueryTable.Refresh.SortState.SortMethod);
        Assert.True(edited.QueryTable.Refresh.SortState.ColumnSort);
        Assert.Equal("A3:B3", edited.QueryTable.Refresh.SortState.Conditions[0].Reference);
        Assert.False(edited.QueryTable.Refresh.SortState.Conditions[0].Descending);
        Assert.Equal("South,North", edited.QueryTable.Refresh.SortState.Conditions[0].CustomList);
        Assert.Equal("A2:B2", edited.QueryTable.Refresh.SortState.Conditions[1].Reference);
        Assert.Equal(1U, edited.QueryTable.Refresh.SortState.Conditions[1].Icon.IconId);
        Assert.Equal(query.Source.QueryPartPath, edited.QueryTable.Source.QueryPartPath);
        Assert.Equal(query.Source.RelationshipId, edited.QueryTable.Source.RelationshipId);
    }

    [Fact]
    public void WorkbookConnectionsRejectFabricationIdentityTopologyAndBindingTamper()
    {
        var sourceFree = TableExportRequest();
        sourceFree.Artifact.Workbook.Connections.Add(new SpreadsheetConnectionArtifact
        {
            ConnectionId = 7,
            Name = "Fabricated",
            Type = 5,
            RefreshedVersion = 8,
        });
        var response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(sourceFree.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_workbook_connection", Assert.Single(response.Diagnostics).Code);

        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(TableExportRequest().ToByteArray()));
        var source = AddQueryTableGraph(authored.File.ToByteArray());

        var imported = Import(source);
        imported.Artifact.Workbook.Connections[0].ConnectionId = 8;
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Contains("id/type/version identity", Assert.Single(response.Diagnostics).Message);

        imported = Import(source);
        imported.Artifact.Workbook.Connections[0].Type = 6;
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Equal("invalid_workbook_connection", Assert.Single(response.Diagnostics).Code);

        imported = Import(source);
        imported.Artifact.Workbook.Connections[0].RefreshedVersion = 256;
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Equal("invalid_workbook_connection", Assert.Single(response.Diagnostics).Code);

        imported = Import(source);
        imported.Artifact.Workbook.Connections.Clear();
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Contains("add or remove recognized workbook connections", Assert.Single(response.Diagnostics).Message);

        imported = Import(source);
        imported.Artifact.Workbook.Connections[0].IntervalMinutes = 32_768;
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Contains("bounded source-editable profile", Assert.Single(response.Diagnostics).Message);

        imported = Import(source);
        imported.Artifact.Workbook.Connections[0].Source.ConnectionXmlSha256 = new string('0', 64);
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Contains("source binding", Assert.Single(response.Diagnostics).Message);
    }

    [Fact]
    public void OpaqueWorkbookConnectionStaysHiddenAndByteExactDuringQueryEdit()
    {
        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(TableExportRequest().ToByteArray()));
        var source = AddQueryTableGraph(authored.File.ToByteArray(), opaqueConnection: true);
        AssertOffice2021Valid(source);
        var connectionXml = ReadEntry(source, "xl/connections.xml");
        var imported = Import(source);
        Assert.Empty(imported.Artifact.Workbook.Connections);
        var query = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Tables).QueryTable;
        Assert.NotNull(query);
        query.Name = "Opaque connection retained";

        var response = Export(imported.Artifact);
        Assert.True(response.Ok, string.Join("\n", response.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(response.File.ToByteArray());
        Assert.Equal(connectionXml, ReadEntry(response.File.ToByteArray(), "xl/connections.xml"));

        imported = Import(source);
        imported.Artifact.Workbook.Connections.Add(new SpreadsheetConnectionArtifact
        {
            ConnectionId = 7,
            Name = "Lossy replacement",
            Type = 5,
            RefreshedVersion = 8,
        });
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Equal("invalid_workbook_connection", Assert.Single(response.Diagnostics).Code);
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

        imported = Import(source);
        imported.Artifact.Workbook.Worksheets[0].Tables[0].QueryTable.Refresh.SortState.ColumnSort = true;
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Contains("row sort condition", Assert.Single(response.Diagnostics).Message);

        imported = Import(source);
        imported.Artifact.Workbook.Worksheets[0].Tables[0].QueryTable.Refresh.SortState.SortMethod = "radical";
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Contains("locale-specific sort method", Assert.Single(response.Diagnostics).Message);

        imported = Import(source);
        imported.Artifact.Workbook.Worksheets[0].Tables[0].QueryTable.Refresh.SortState.Conditions[1].CustomList = "up,flat,down";
        response = Export(imported.Artifact);
        Assert.False(response.Ok);
        Assert.Contains("custom-list value-sort", Assert.Single(response.Diagnostics).Message);
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
        Assert.Contains("columnSort=\"1\"", queryXml);
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
    public void ProtocolAuthorsImportsAndEditsWorksheetColumnSortState()
    {
        var request = ExportRequest();
        var sort = new SpreadsheetTableSortStateArtifact
        {
            Reference = "A1:B2", CaseSensitive = true, SortMethod = "stroke", ColumnSort = true,
        };
        sort.Conditions.Add(new SpreadsheetTableSortConditionArtifact { Reference = "A2:B2", Descending = true, CustomList = "Q2,Q1" });
        sort.Conditions.Add(new SpreadsheetTableSortConditionArtifact { Reference = "A1:B1" });
        request.Artifact.Workbook.Worksheets[0].SortState = sort;

        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var xml = System.Text.Encoding.UTF8.GetString(ReadEntry(exported.File.ToByteArray(), "xl/worksheets/sheet1.xml"));
        Assert.Contains("<x:sortState", xml);
        Assert.Contains("ref=\"A1:B2\"", xml);
        Assert.Contains("caseSensitive=\"1\"", xml);
        Assert.Contains("sortMethod=\"stroke\"", xml);
        Assert.Contains("columnSort=\"1\"", xml);
        Assert.Contains("ref=\"A2:B2\"", xml);
        Assert.Contains("descending=\"1\"", xml);
        Assert.Contains("customList=\"Q2,Q1\"", xml);
        Assert.Contains("ref=\"A1:B1\"", xml);
        AssertOffice2021Valid(exported.File.ToByteArray());

        var imported = Import(exported.File.ToByteArray());
        var importedSort = Assert.IsType<SpreadsheetTableSortStateArtifact>(imported.Artifact.Workbook.Worksheets[0].SortState);
        Assert.True(importedSort.HasColumnSort);
        Assert.True(importedSort.ColumnSort);
        Assert.Equal("A2:B2", importedSort.Conditions[0].Reference);

        importedSort.ColumnSort = false;
        importedSort.Conditions[0].Reference = "B1:B2";
        importedSort.Conditions[1].Reference = "A1:A2";
        var edited = Export(imported.Artifact);
        Assert.True(edited.Ok, string.Join("\n", edited.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        xml = System.Text.Encoding.UTF8.GetString(ReadEntry(edited.File.ToByteArray(), "xl/worksheets/sheet1.xml"));
        Assert.Contains("columnSort=\"0\"", xml);
        var reimported = Assert.IsType<SpreadsheetTableSortStateArtifact>(Import(edited.File.ToByteArray()).Artifact.Workbook.Worksheets[0].SortState);
        Assert.True(reimported.HasColumnSort);
        Assert.False(reimported.ColumnSort);
        Assert.Equal("B1:B2", reimported.Conditions[0].Reference);

        request = ExportRequest();
        request.Artifact.Workbook.Worksheets[0].SortState = new SpreadsheetTableSortStateArtifact
        {
            Reference = "A1:B2", ColumnSort = true,
        };
        request.Artifact.Workbook.Worksheets[0].SortState.Conditions.Add(new SpreadsheetTableSortConditionArtifact { Reference = "A1:A2" });
        var rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_worksheet_sort", Assert.Single(rejected.Diagnostics).Code);
        Assert.Contains("row sort condition", Assert.Single(rejected.Diagnostics).Message);
    }

    [Fact]
    public void UnsupportedWorksheetColumnSortGeometryStaysHiddenAndPreserved()
    {
        var request = ExportRequest();
        var sort = new SpreadsheetTableSortStateArtifact { Reference = "A1:B2", ColumnSort = true };
        sort.Conditions.Add(new SpreadsheetTableSortConditionArtifact { Reference = "A2:B2", Descending = true });
        sort.Conditions.Add(new SpreadsheetTableSortConditionArtifact { Reference = "A1:B1" });
        request.Artifact.Workbook.Worksheets[0].SortState = sort;
        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        var source = SetWorksheetSortConditionReferences(authored.File.ToByteArray(), "A1:A2", "B1:B2");
        AssertOffice2021Valid(source);

        var imported = Import(source);
        Assert.Null(imported.Artifact.Workbook.Worksheets[0].SortState);
        imported.Artifact.Workbook.Worksheets[0].Cells.Single(cell => cell.Row == 0 && cell.Column == 1).NumberValue = 43;
        var preserved = Export(imported.Artifact);
        Assert.True(preserved.Ok, string.Join("\n", preserved.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal(ReadWorksheetSortXml(source), ReadWorksheetSortXml(preserved.File.ToByteArray()));

        imported = Import(source);
        imported.Artifact.Workbook.Worksheets[0].SortState = sort.Clone();
        var rejected = Export(imported.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_worksheet_sort", Assert.Single(rejected.Diagnostics).Code);
        Assert.Contains("unsupported native sortState", Assert.Single(rejected.Diagnostics).Message);
    }

    [Fact]
    public void ProtocolAuthorsAndImportsWorksheetTableValueSortState()
    {
        var exported = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(SortTableExportRequest().ToByteArray()));
        Assert.True(exported.Ok, string.Join("\n", exported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var xml = System.Text.Encoding.UTF8.GetString(ReadEntry(exported.File.ToByteArray(), "xl/tables/table1.xml"));
        Assert.Contains("<x:sortState ref=\"A2:B3\" caseSensitive=\"1\" sortMethod=\"stroke\"><x:sortCondition ref=\"B2:B3\" descending=\"1\" customList=\"High,Medium,Low\" /><x:sortCondition ref=\"A2:A3\" /></x:sortState>", xml);
        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));

        var imported = Import(exported.File.ToByteArray());
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var sort = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Tables).SortState;
        Assert.Equal("A2:B3", sort.Reference);
        Assert.True(sort.CaseSensitive);
        Assert.Equal("stroke", sort.SortMethod);
        Assert.Equal(2, sort.Conditions.Count);
        Assert.Equal("B2:B3", sort.Conditions[0].Reference);
        Assert.True(sort.Conditions[0].Descending);
        Assert.Equal("High,Medium,Low", sort.Conditions[0].CustomList);
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
        table.SortState.SortMethod = "none";
        table.SortState.Conditions[0].Descending = false;
        table.SortState.Conditions[0].CustomList = "Low,Medium,High";
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
        Assert.Equal("none", edited.SortState.SortMethod);
        Assert.False(edited.SortState.Conditions[0].Descending);
        Assert.Equal("Low,Medium,High", edited.SortState.Conditions[0].CustomList);
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

        request = SortTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].SortState.ColumnSort = false;
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Contains("inside an AutoFilter", Assert.Single(response.Diagnostics).Message);

        request = SortTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].SortState.SortMethod = "radical";
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Contains("locale-specific sort method", Assert.Single(response.Diagnostics).Message);

        request = SortTableExportRequest();
        request.Artifact.Workbook.Worksheets[0].Tables[0].SortState.Conditions[1].Icon = new SpreadsheetTableIconArtifact { IconSet = "3Arrows" };
        request.Artifact.Workbook.Worksheets[0].Tables[0].SortState.Conditions[1].CustomList = "up,flat,down";
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Contains("custom-list value-sort", Assert.Single(response.Diagnostics).Message);
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
    public void ProtocolAuthorsImportsAndSourcePreservesDynamicArrayMetadata()
    {
        var request = FormulaExportRequest();
        request.Artifact.Workbook.Worksheets[0].Cells.Add(new CellArtifact
        {
            Row = 0,
            Column = 6,
            Formula = "=SEQUENCE(2,2)",
            NumberValue = 1,
            FormulaMetadata = new CellFormulaMetadata { Kind = CellFormulaKind.DynamicArray, Reference = "G1:H2" },
        });
        request.Artifact.Workbook.Worksheets[0].Cells.Add(new CellArtifact { Row = 0, Column = 7, NumberValue = 2 });
        request.Artifact.Workbook.Worksheets[0].Cells.Add(new CellArtifact { Row = 1, Column = 6, NumberValue = 3 });
        request.Artifact.Workbook.Worksheets[0].Cells.Add(new CellArtifact { Row = 1, Column = 7, NumberValue = 4 });

        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.True(authored.Ok, string.Join("\n", authored.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(authored.File.ToByteArray());
        string metadataPartPath;
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            var workbookPart = document.WorkbookPart!;
            var metadata = workbookPart.CellMetadataPart?.Metadata;
            Assert.NotNull(metadata);
            metadataPartPath = workbookPart.CellMetadataPart!.Uri.OriginalString.TrimStart('/');
            Assert.Contains("dynamicArrayProperties", metadata!.OuterXml);
            Assert.Contains("fDynamic=\"1\"", metadata.OuterXml);
            Assert.Contains("fCollapsed=\"0\"", metadata.OuterXml);
            var cells = workbookPart.WorksheetParts.Single().Worksheet!.Descendants<Cell>().ToDictionary(item => item.CellReference!.Value!);
            Assert.Equal(1U, cells["G1"].CellMetaIndex!.Value);
            Assert.Equal(CellFormulaValues.Array, cells["G1"].CellFormula!.FormulaType!.Value);
            Assert.Equal("G1:H2", cells["G1"].CellFormula!.Reference!.Value);
            Assert.Null(cells["E1"].CellMetaIndex);
        }

        var imported = Import(authored.File.ToByteArray());
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var dynamic = imported.Artifact.Workbook.Worksheets[0].Cells.Single(cell => cell.Row == 0 && cell.Column == 6);
        Assert.Equal(CellFormulaKind.DynamicArray, dynamic.FormulaMetadata.Kind);
        Assert.Equal("G1:H2", dynamic.FormulaMetadata.Reference);

        var metadataBefore = ReadEntry(authored.File.ToByteArray(), metadataPartPath);
        dynamic.Formula = "=SEQUENCE(2,2,10,1)";
        dynamic.NumberValue = 10;
        var edited = Export(imported.Artifact);
        Assert.True(edited.Ok, string.Join("\n", edited.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal(metadataBefore, ReadEntry(edited.File.ToByteArray(), metadataPartPath));
        using (var stream = new MemoryStream(edited.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            var anchor = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "G1");
            Assert.Equal(1U, anchor.CellMetaIndex!.Value);
            Assert.Equal("SEQUENCE(2,2,10,1)", anchor.CellFormula!.Text);
        }

        dynamic.FormulaMetadata = null;
        var detached = Export(imported.Artifact);
        Assert.True(detached.Ok, string.Join("\n", detached.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        using var detachedStream = new MemoryStream(detached.File.ToByteArray());
        using var detachedDocument = SpreadsheetDocument.Open(detachedStream, false);
        var detachedAnchor = detachedDocument.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "G1");
        Assert.Null(detachedAnchor.CellMetaIndex);
        Assert.Null(detachedAnchor.CellFormula!.FormulaType);
        Assert.Equal(metadataBefore, ReadEntry(detached.File.ToByteArray(), metadataPartPath));
    }

    [Fact]
    public void DynamicArrayMetadataEditsFailClosedWithoutAnOwnedRecord()
    {
        var first = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(FormulaExportRequest().ToByteArray()));
        var imported = Import(first.File.ToByteArray());
        imported.Artifact.Workbook.Worksheets[0].Cells.Add(new CellArtifact
        {
            Row = 0,
            Column = 6,
            Formula = "=SEQUENCE(2)",
            NumberValue = 1,
            FormulaMetadata = new CellFormulaMetadata { Kind = CellFormulaKind.DynamicArray, Reference = "G1:G2" },
        });
        var rejected = Export(imported.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_dynamic_array_edit", Assert.Single(rejected.Diagnostics).Code);

        var dynamicRequest = FormulaExportRequest();
        dynamicRequest.Artifact.Workbook.Worksheets[0].Cells.Add(new CellArtifact
        {
            Row = 0,
            Column = 6,
            Formula = "=SEQUENCE(2)",
            NumberValue = 1,
            FormulaMetadata = new CellFormulaMetadata { Kind = CellFormulaKind.DynamicArray, Reference = "G1:G2" },
        });
        var dynamicFile = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(dynamicRequest.ToByteArray()));
        var malformed = SetCellMetadataIndex(dynamicFile.File.ToByteArray(), "C1", 1U);
        var malformedImport = Import(malformed);
        Assert.False(malformedImport.Ok);
        Assert.Equal("invalid_cell_formula", Assert.Single(malformedImport.Diagnostics).Code);
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

    [Fact]
    public void ProtocolAuthorsImportsAndSourcePreservesWorksheetPictures()
    {
        var request = PictureExportRequest();
        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.True(authored.Ok, string.Join("\n", authored.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(authored.File.ToByteArray());
        AssertPicture(authored.File.ToByteArray(), "Quarter mark", "Quarterly performance", 3, 2, 1_143_000L, 762_000L, residual: false);

        var source = AddPictureResidual(authored.File.ToByteArray());
        var imported = Import(source);
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var asset = Assert.Single(imported.Artifact.Assets);
        Assert.StartsWith("asset/workbook/image/", asset.Id, StringComparison.Ordinal);
        var image = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Images);
        Assert.Equal("Quarter mark", image.Name);
        Assert.Equal("Quarterly performance", image.AltText);
        Assert.Equal(asset.Id, image.AssetId);
        Assert.Equal(3U, image.Anchor.Row);
        Assert.Equal(2U, image.Anchor.Column);
        Assert.True(image.Source.Editable);
        Assert.Equal(64, image.Source.DrawingXmlSha256.Length);
        Assert.Equal(64, image.Source.AnchorXmlSha256.Length);
        Assert.Equal(64, image.Source.SemanticSha256.Length);

        image.Name = "Updated quarter mark";
        image.AltText = "Updated quarterly performance";
        image.Anchor.Row = 5;
        image.Anchor.Column = 4;
        image.Anchor.WidthEmu = 1_524_000L;
        image.Anchor.HeightEmu = 952_500L;
        var preserved = Export(imported.Artifact);
        Assert.True(preserved.Ok, string.Join("\n", preserved.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(preserved.File.ToByteArray());
        AssertPicture(preserved.File.ToByteArray(), "Updated quarter mark", "Updated quarterly performance", 5, 4, 1_524_000L, 952_500L, residual: true);
        var reimported = Import(preserved.File.ToByteArray());
        Assert.True(reimported.Ok, string.Join("\n", reimported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal("Updated quarter mark", Assert.Single(reimported.Artifact.Workbook.Worksheets[0].Images).Name);
        Assert.Equal(asset.Data, Assert.Single(reimported.Artifact.Assets).Data);

        var removed = Import(source);
        removed.Artifact.Workbook.Worksheets[0].Images.Clear();
        var rejected = Export(removed.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_image_topology", Assert.Single(rejected.Diagnostics).Code);

        var rebound = Import(source);
        var replacementBytes = rebound.Artifact.Assets[0].Data.ToByteArray();
        replacementBytes[^1] ^= 1;
        var replacement = WorksheetImageAsset(replacementBytes, "image/png", "replacement.png");
        rebound.Artifact.Assets.Add(replacement);
        rebound.Artifact.Workbook.Worksheets[0].Images[0].AssetId = replacement.Id;
        var replaced = Export(rebound.Artifact);
        Assert.True(replaced.Ok, string.Join("\n", replaced.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(replaced.File.ToByteArray());
        Assert.Equal(ReadPicturePartPath(source), ReadPicturePartPath(replaced.File.ToByteArray()));
        Assert.Equal(replacementBytes, ReadPictureBytes(replaced.File.ToByteArray()));
        var replacedImport = Import(replaced.File.ToByteArray());
        Assert.True(replacedImport.Ok, string.Join("\n", replacedImport.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal(replacement.Data, Assert.Single(replacedImport.Artifact.Assets).Data);

        var tampered = Import(source);
        tampered.Artifact.Workbook.Worksheets[0].Images[0].Source.AnchorXmlSha256 = new string('0', 64);
        rejected = Export(tampered.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("spreadsheet_image_source_binding_mismatch", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void ProtocolAuthorsImportsAndSourcePreservesWorksheetChartsBesidePictures()
    {
        var request = ChartExportRequest();
        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.True(authored.Ok, string.Join("\n", authored.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(authored.File.ToByteArray());
        AssertChartAxes(authored.File.ToByteArray(), "Quarter", "@", 2, "Revenue", "$#,##0.0", 0, 100, 25);
        AssertChartTextStyles(authored.File.ToByteArray(), 12, 10, 9);
        AssertChartSeriesFill(authored.File.ToByteArray(), "F472B6");
        AssertChartSeriesLine(authored.File.ToByteArray(), "0EA5E9", "dash", 2);
        AssertChartSeriesMarker(authored.File.ToByteArray(), "diamond", 8);
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            var drawingPart = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!;
            Assert.Single(drawingPart.WorksheetDrawing!.Descendants<Xdr.Picture>());
            Assert.Single(drawingPart.WorksheetDrawing.Descendants<Xdr.GraphicFrame>());
            Assert.Single(drawingPart.ChartParts);
            Assert.Contains("Quarter trend", ReadPartText(drawingPart.ChartParts.Single()), StringComparison.Ordinal);
        }

        var source = AddChartResidual(authored.File.ToByteArray());
        var imported = Import(source);
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var sheet = Assert.Single(imported.Artifact.Workbook.Worksheets);
        Assert.Single(sheet.Images);
        var chart = Assert.Single(sheet.Charts);
        Assert.Equal(SpreadsheetChartType.Line, chart.Type);
        Assert.Equal("Quarter trend", chart.Title);
        Assert.Equal(12, chart.TitleTextStyle.FontSizePoints);
        Assert.True(chart.HasLegend);
        Assert.Equal(["Q1", "Q2"], chart.Categories);
        Assert.Equal([42.5, 85], Assert.Single(chart.Series).Values);
        Assert.Equal(SpreadsheetColor.SourceOneofCase.Rgb, chart.Series[0].Fill.SourceCase);
        Assert.Equal("F472B6", chart.Series[0].Fill.Rgb);
        Assert.Equal("0EA5E9", chart.Series[0].Line.Color.Rgb);
        Assert.Equal(SpreadsheetChartLineDashStyle.Dashed, chart.Series[0].Line.DashStyle);
        Assert.True(chart.Series[0].Line.HasWidthPoints);
        Assert.Equal(2, chart.Series[0].Line.WidthPoints);
        Assert.Equal(SpreadsheetChartMarkerSymbol.Diamond, chart.Series[0].Marker.Symbol);
        Assert.True(chart.Series[0].Marker.HasSize);
        Assert.Equal(8U, chart.Series[0].Marker.Size);
        Assert.Equal("'Summary'!$A$1:$A$2", chart.Series[0].CategoryFormula);
        Assert.Equal("'Summary'!$B$1:$B$2", chart.Series[0].ValueFormula);
        Assert.Equal("Quarter", chart.XAxis.Title);
        Assert.Equal("@", chart.XAxis.NumberFormatCode);
        Assert.True(chart.XAxis.HasTickLabelInterval);
        Assert.Equal(2U, chart.XAxis.TickLabelInterval);
        Assert.Equal(10, chart.XAxis.TextStyle.FontSizePoints);
        Assert.Equal("Revenue", chart.YAxis.Title);
        Assert.Equal("$#,##0.0", chart.YAxis.NumberFormatCode);
        Assert.True(chart.YAxis.HasMinimum);
        Assert.Equal(0, chart.YAxis.Minimum);
        Assert.True(chart.YAxis.HasMaximum);
        Assert.Equal(100, chart.YAxis.Maximum);
        Assert.True(chart.YAxis.HasMajorUnit);
        Assert.Equal(25, chart.YAxis.MajorUnit);
        Assert.Equal(9, chart.YAxis.TextStyle.FontSizePoints);
        Assert.True(chart.Source.Editable);
        Assert.Equal(64, chart.Source.DrawingXmlSha256.Length);
        Assert.Equal(64, chart.Source.ChartXmlSha256.Length);
        Assert.Equal(64, chart.Source.SemanticSha256.Length);

        sheet.Images[0].Name = "Picture edited with chart";
        chart.Name = "Edited native chart";
        chart.Title = "Edited quarter trend";
        chart.TitleTextStyle.FontSizePoints = 15;
        chart.HasLegend = false;
        chart.Categories[1] = "Q2 actual";
        chart.Series[0].Name = "Actual revenue";
        chart.Series[0].Values[1] = 90;
        chart.Series[0].Fill.Rgb = "2563EB";
        chart.Series[0].Line.Color.Rgb = "7C3AED";
        chart.Series[0].Line.DashStyle = SpreadsheetChartLineDashStyle.DashDot;
        chart.Series[0].Line.WidthPoints = 2.5;
        chart.Series[0].Marker.Symbol = SpreadsheetChartMarkerSymbol.Triangle;
        chart.Series[0].Marker.Size = 10;
        chart.XAxis.Title = "Fiscal quarter";
        chart.XAxis.NumberFormatCode = "mmm";
        chart.XAxis.TickLabelInterval = 1;
        chart.XAxis.TextStyle.FontSizePoints = 11;
        chart.YAxis.Title = "Revenue USD";
        chart.YAxis.NumberFormatCode = "$0";
        chart.YAxis.Minimum = -10;
        chart.YAxis.Maximum = 120;
        chart.YAxis.MajorUnit = 10;
        chart.YAxis.TextStyle = null;
        var preserved = Export(imported.Artifact);
        Assert.True(preserved.Ok, string.Join("\n", preserved.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(preserved.File.ToByteArray());
        AssertChartAxes(preserved.File.ToByteArray(), "Fiscal quarter", "mmm", 1, "Revenue USD", "$0", -10, 120, 10);
        AssertChartTextStyles(preserved.File.ToByteArray(), 15, 11, null);
        AssertChartSeriesFill(preserved.File.ToByteArray(), "2563EB");
        AssertChartSeriesLine(preserved.File.ToByteArray(), "7C3AED", "dashDot", 2.5);
        AssertChartSeriesMarker(preserved.File.ToByteArray(), "triangle", 10);
        using (var stream = new MemoryStream(preserved.File.ToByteArray()))
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            var drawingPart = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!;
            Assert.Equal("Picture edited with chart", drawingPart.WorksheetDrawing!.Descendants<Xdr.Picture>().Single().NonVisualPictureProperties!.NonVisualDrawingProperties!.Name!.Value);
            Assert.Equal("Edited native chart", drawingPart.WorksheetDrawing.Descendants<Xdr.GraphicFrame>().Single().NonVisualGraphicFrameProperties!.NonVisualDrawingProperties!.Name!.Value);
            var xml = ReadPartText(drawingPart.ChartParts.Single());
            Assert.Contains("Edited quarter trend", xml, StringComparison.Ordinal);
            Assert.Contains("Q2 actual", xml, StringComparison.Ordinal);
            Assert.Contains(">90<", xml, StringComparison.Ordinal);
            Assert.Contains("preserve-me", xml, StringComparison.Ordinal);
            Assert.DoesNotContain("<c:legend>", xml, StringComparison.Ordinal);
        }

        var removedFill = Import(preserved.File.ToByteArray());
        removedFill.Artifact.Workbook.Worksheets[0].Charts[0].Series[0].Fill = null;
        var withoutFill = Export(removedFill.Artifact);
        Assert.True(withoutFill.Ok, string.Join("\n", withoutFill.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertChartSeriesFill(withoutFill.File.ToByteArray(), null);
        var addedFill = Import(withoutFill.File.ToByteArray());
        Assert.Null(addedFill.Artifact.Workbook.Worksheets[0].Charts[0].Series[0].Fill);
        addedFill.Artifact.Workbook.Worksheets[0].Charts[0].Series[0].Fill = new SpreadsheetColor { Rgb = "22C55E" };
        var withAddedFill = Export(addedFill.Artifact);
        Assert.True(withAddedFill.Ok, string.Join("\n", withAddedFill.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertChartSeriesFill(withAddedFill.File.ToByteArray(), "22C55E");

        var removedLine = Import(preserved.File.ToByteArray());
        removedLine.Artifact.Workbook.Worksheets[0].Charts[0].Series[0].Line = null;
        var withoutLine = Export(removedLine.Artifact);
        Assert.True(withoutLine.Ok, string.Join("\n", withoutLine.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertChartSeriesLine(withoutLine.File.ToByteArray(), null, null, null);
        var addedLine = Import(withoutLine.File.ToByteArray());
        Assert.Null(addedLine.Artifact.Workbook.Worksheets[0].Charts[0].Series[0].Line);
        addedLine.Artifact.Workbook.Worksheets[0].Charts[0].Series[0].Line = new SpreadsheetChartLineStyleArtifact
        {
            Color = new SpreadsheetColor { Rgb = "22C55E" },
            DashStyle = SpreadsheetChartLineDashStyle.Dotted,
            WidthPoints = 1.25,
        };
        var withAddedLine = Export(addedLine.Artifact);
        Assert.True(withAddedLine.Ok, string.Join("\n", withAddedLine.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertChartSeriesLine(withAddedLine.File.ToByteArray(), "22C55E", "dot", 1.25);

        var removedMarker = Import(preserved.File.ToByteArray());
        removedMarker.Artifact.Workbook.Worksheets[0].Charts[0].Series[0].Marker = null;
        var withoutMarker = Export(removedMarker.Artifact);
        Assert.True(withoutMarker.Ok, string.Join("\n", withoutMarker.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertChartSeriesMarker(withoutMarker.File.ToByteArray(), null, null);
        var addedMarker = Import(withoutMarker.File.ToByteArray());
        Assert.Null(addedMarker.Artifact.Workbook.Worksheets[0].Charts[0].Series[0].Marker);
        addedMarker.Artifact.Workbook.Worksheets[0].Charts[0].Series[0].Marker = new SpreadsheetChartMarkerArtifact { Symbol = SpreadsheetChartMarkerSymbol.Plus, Size = 12 };
        var withAddedMarker = Export(addedMarker.Artifact);
        Assert.True(withAddedMarker.Ok, string.Join("\n", withAddedMarker.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertChartSeriesMarker(withAddedMarker.File.ToByteArray(), "plus", 12);

        var addedTextStyle = Import(preserved.File.ToByteArray());
        addedTextStyle.Artifact.Workbook.Worksheets[0].Charts[0].YAxis.TextStyle = new SpreadsheetChartTextStyleArtifact { FontSizePoints = 8.5 };
        var withAddedTextStyle = Export(addedTextStyle.Artifact);
        Assert.True(withAddedTextStyle.Ok, string.Join("\n", withAddedTextStyle.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(withAddedTextStyle.File.ToByteArray());
        AssertChartTextStyles(withAddedTextStyle.File.ToByteArray(), 15, 11, 8.5);

        var removed = Import(source);
        removed.Artifact.Workbook.Worksheets[0].Charts.Clear();
        var rejected = Export(removed.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_chart_topology", Assert.Single(rejected.Diagnostics).Code);

        var changedType = Import(source);
        changedType.Artifact.Workbook.Worksheets[0].Charts[0].Series[0].Marker = null;
        changedType.Artifact.Workbook.Worksheets[0].Charts[0].Type = SpreadsheetChartType.Bar;
        rejected = Export(changedType.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_spreadsheet_chart_edit", Assert.Single(rejected.Diagnostics).Code);

        var referencedSource = SetChartSeriesNameReference(authored.File.ToByteArray());
        var referencedXml = ReadChartXml(referencedSource);
        var referenced = Import(referencedSource);
        Assert.True(referenced.Ok, string.Join("\n", referenced.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var readOnlyChart = Assert.Single(referenced.Artifact.Workbook.Worksheets[0].Charts);
        Assert.Equal("Revenue from cache", Assert.Single(readOnlyChart.Series).Name);
        Assert.False(readOnlyChart.Source.Editable);
        var exactRoundTrip = Export(referenced.Artifact);
        Assert.True(exactRoundTrip.Ok, string.Join("\n", exactRoundTrip.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal(referencedXml, ReadChartXml(exactRoundTrip.File.ToByteArray()));
        readOnlyChart.Title = "Forbidden edit";
        rejected = Export(referenced.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_spreadsheet_chart_edit", Assert.Single(rejected.Diagnostics).Code);

        var logarithmicSource = SetChartValueAxisLogarithmic(authored.File.ToByteArray());
        var logarithmicXml = ReadChartXml(logarithmicSource);
        var logarithmic = Import(logarithmicSource);
        Assert.True(logarithmic.Ok, string.Join("\n", logarithmic.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var logarithmicChart = Assert.Single(logarithmic.Artifact.Workbook.Worksheets[0].Charts);
        Assert.False(logarithmicChart.Source.Editable);
        var logarithmicRoundTrip = Export(logarithmic.Artifact);
        Assert.True(logarithmicRoundTrip.Ok, string.Join("\n", logarithmicRoundTrip.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal(logarithmicXml, ReadChartXml(logarithmicRoundTrip.File.ToByteArray()));
        logarithmicChart.YAxis.Title = "Forbidden logarithmic edit";
        rejected = Export(logarithmic.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_spreadsheet_chart_edit", Assert.Single(rejected.Diagnostics).Code);

        var schemeFillSource = SetChartSeriesSchemeFill(authored.File.ToByteArray());
        var schemeFillXml = ReadChartXml(schemeFillSource);
        var schemeFill = Import(schemeFillSource);
        Assert.True(schemeFill.Ok, string.Join("\n", schemeFill.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var schemeFillChart = Assert.Single(schemeFill.Artifact.Workbook.Worksheets[0].Charts);
        Assert.Null(Assert.Single(schemeFillChart.Series).Fill);
        Assert.False(schemeFillChart.Source.Editable);
        var schemeFillRoundTrip = Export(schemeFill.Artifact);
        Assert.True(schemeFillRoundTrip.Ok, string.Join("\n", schemeFillRoundTrip.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal(schemeFillXml, ReadChartXml(schemeFillRoundTrip.File.ToByteArray()));
        schemeFillChart.Series[0].Fill = new SpreadsheetColor { Rgb = "E11D48" };
        rejected = Export(schemeFill.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_spreadsheet_chart_edit", Assert.Single(rejected.Diagnostics).Code);

        var complexLineSource = SetChartSeriesComplexLine(authored.File.ToByteArray());
        var complexLineXml = ReadChartXml(complexLineSource);
        var complexLine = Import(complexLineSource);
        Assert.True(complexLine.Ok, string.Join("\n", complexLine.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var complexLineChart = Assert.Single(complexLine.Artifact.Workbook.Worksheets[0].Charts);
        Assert.Null(Assert.Single(complexLineChart.Series).Line);
        Assert.False(complexLineChart.Source.Editable);
        var complexLineRoundTrip = Export(complexLine.Artifact);
        Assert.True(complexLineRoundTrip.Ok, string.Join("\n", complexLineRoundTrip.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal(complexLineXml, ReadChartXml(complexLineRoundTrip.File.ToByteArray()));
        complexLineChart.Series[0].Line = new SpreadsheetChartLineStyleArtifact { Color = new SpreadsheetColor { Rgb = "E11D48" } };
        rejected = Export(complexLine.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_spreadsheet_chart_edit", Assert.Single(rejected.Diagnostics).Code);

        var complexMarkerSource = SetChartSeriesComplexMarker(authored.File.ToByteArray());
        var complexMarkerXml = ReadChartXml(complexMarkerSource);
        var complexMarker = Import(complexMarkerSource);
        Assert.True(complexMarker.Ok, string.Join("\n", complexMarker.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var complexMarkerChart = Assert.Single(complexMarker.Artifact.Workbook.Worksheets[0].Charts);
        Assert.Null(Assert.Single(complexMarkerChart.Series).Marker);
        Assert.False(complexMarkerChart.Source.Editable);
        var complexMarkerRoundTrip = Export(complexMarker.Artifact);
        Assert.True(complexMarkerRoundTrip.Ok, string.Join("\n", complexMarkerRoundTrip.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal(complexMarkerXml, ReadChartXml(complexMarkerRoundTrip.File.ToByteArray()));
        complexMarkerChart.Series[0].Marker = new SpreadsheetChartMarkerArtifact { Symbol = SpreadsheetChartMarkerSymbol.Circle };
        rejected = Export(complexMarker.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_spreadsheet_chart_edit", Assert.Single(rejected.Diagnostics).Code);

        var complexTextSource = SetChartAxisComplexTextStyle(authored.File.ToByteArray());
        var complexTextXml = ReadChartXml(complexTextSource);
        var complexText = Import(complexTextSource);
        Assert.True(complexText.Ok, string.Join("\n", complexText.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var complexTextChart = Assert.Single(complexText.Artifact.Workbook.Worksheets[0].Charts);
        Assert.False(complexTextChart.Source.Editable);
        Assert.Null(complexTextChart.XAxis.TextStyle);
        var complexTextRoundTrip = Export(complexText.Artifact);
        Assert.True(complexTextRoundTrip.Ok, string.Join("\n", complexTextRoundTrip.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal(complexTextXml, ReadChartXml(complexTextRoundTrip.File.ToByteArray()));
        complexTextChart.XAxis.TextStyle = new SpreadsheetChartTextStyleArtifact { FontSizePoints = 14 };
        rejected = Export(complexText.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_spreadsheet_chart_edit", Assert.Single(rejected.Diagnostics).Code);

        var tampered = Import(source);
        tampered.Artifact.Workbook.Worksheets[0].Charts[0].Source.ChartXmlSha256 = new string('0', 64);
        rejected = Export(tampered.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("spreadsheet_chart_source_binding_mismatch", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void ProtocolRejectsInvalidWorksheetChartAxes()
    {
        var reversed = ChartExportRequest();
        reversed.Artifact.Workbook.Worksheets[0].Charts[0].YAxis.Minimum = 100;
        reversed.Artifact.Workbook.Worksheets[0].Charts[0].YAxis.Maximum = 10;
        var rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(reversed.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_chart", Assert.Single(rejected.Diagnostics).Code);

        var oversizedFont = ChartExportRequest();
        oversizedFont.Artifact.Workbook.Worksheets[0].Charts[0].TitleTextStyle.FontSizePoints = 4_001;
        rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(oversizedFont.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_chart", Assert.Single(rejected.Diagnostics).Code);

        var emptyTitle = ChartExportRequest();
        emptyTitle.Artifact.Workbook.Worksheets[0].Charts[0].Title = string.Empty;
        rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(emptyTitle.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_chart", Assert.Single(rejected.Diagnostics).Code);

        var pie = ChartExportRequest();
        pie.Artifact.Workbook.Worksheets[0].Charts[0].Type = SpreadsheetChartType.Pie;
        rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(pie.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_chart", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void ProtocolRejectsInvalidWorksheetChartSeriesFills()
    {
        var theme = ChartExportRequest();
        theme.Artifact.Workbook.Worksheets[0].Charts[0].Series[0].Fill = new SpreadsheetColor { Theme = 4 };
        var rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(theme.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_chart", Assert.Single(rejected.Diagnostics).Code);

        var tinted = ChartExportRequest();
        tinted.Artifact.Workbook.Worksheets[0].Charts[0].Series[0].Fill = new SpreadsheetColor { Rgb = "2563EB", Tint = 0 };
        rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(tinted.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_chart", Assert.Single(rejected.Diagnostics).Code);

        var malformed = ChartExportRequest();
        malformed.Artifact.Workbook.Worksheets[0].Charts[0].Series[0].Fill = new SpreadsheetColor { Rgb = "blue" };
        rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(malformed.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_chart", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void ProtocolRejectsInvalidWorksheetChartSeriesLines()
    {
        var theme = ChartExportRequest();
        theme.Artifact.Workbook.Worksheets[0].Charts[0].Series[0].Line.Color = new SpreadsheetColor { Theme = 4 };
        var rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(theme.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_chart", Assert.Single(rejected.Diagnostics).Code);

        var tinted = ChartExportRequest();
        tinted.Artifact.Workbook.Worksheets[0].Charts[0].Series[0].Line.Color.Tint = 0;
        rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(tinted.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_chart", Assert.Single(rejected.Diagnostics).Code);

        var oversized = ChartExportRequest();
        oversized.Artifact.Workbook.Worksheets[0].Charts[0].Series[0].Line.WidthPoints = 1_584.1;
        rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(oversized.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_chart", Assert.Single(rejected.Diagnostics).Code);

        var unknownDash = ChartExportRequest();
        unknownDash.Artifact.Workbook.Worksheets[0].Charts[0].Series[0].Line.DashStyle = (SpreadsheetChartLineDashStyle)99;
        rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(unknownDash.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_chart", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void ProtocolRejectsInvalidWorksheetChartSeriesMarkers()
    {
        var bar = ChartExportRequest();
        bar.Artifact.Workbook.Worksheets[0].Charts[0].Type = SpreadsheetChartType.Bar;
        var rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(bar.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_chart", Assert.Single(rejected.Diagnostics).Code);

        var tooSmall = ChartExportRequest();
        tooSmall.Artifact.Workbook.Worksheets[0].Charts[0].Series[0].Marker.Size = 1;
        rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(tooSmall.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_chart", Assert.Single(rejected.Diagnostics).Code);

        var tooLarge = ChartExportRequest();
        tooLarge.Artifact.Workbook.Worksheets[0].Charts[0].Series[0].Marker.Size = 73;
        rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(tooLarge.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_chart", Assert.Single(rejected.Diagnostics).Code);

        var unknown = ChartExportRequest();
        unknown.Artifact.Workbook.Worksheets[0].Charts[0].Series[0].Marker.Symbol = (SpreadsheetChartMarkerSymbol)99;
        rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(unknown.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_chart", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void ProtocolAuthorsImportsAndSourcePreservesTwoCellWorksheetPictures()
    {
        var request = TwoCellPictureExportRequest();
        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.True(authored.Ok, string.Join("\n", authored.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(authored.File.ToByteArray());
        AssertTwoCellPicture(authored.File.ToByteArray(), "Quarter mark", "Quarterly performance", 3, 2, 7, 6, Xdr.EditAsValues.OneCell, residual: false);

        var source = AddPictureResidual(authored.File.ToByteArray());
        var imported = Import(source);
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var image = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Images);
        Assert.Null(image.Anchor);
        Assert.NotNull(image.TwoCellAnchor);
        Assert.Equal(3U, image.TwoCellAnchor.From.Row);
        Assert.Equal(2U, image.TwoCellAnchor.From.Column);
        Assert.Equal(7U, image.TwoCellAnchor.To.Row);
        Assert.Equal(6U, image.TwoCellAnchor.To.Column);
        Assert.True(image.TwoCellAnchor.HasEditAs);
        Assert.Equal(SpreadsheetTwoCellEditAs.OneCell, image.TwoCellAnchor.EditAs);

        image.Name = "Updated two-cell mark";
        image.AltText = "Updated two-cell performance";
        image.TwoCellAnchor.From.Row = 4;
        image.TwoCellAnchor.From.Column = 3;
        image.TwoCellAnchor.To.Row = 9;
        image.TwoCellAnchor.To.Column = 8;
        image.TwoCellAnchor.EditAs = SpreadsheetTwoCellEditAs.Absolute;
        var preserved = Export(imported.Artifact);
        Assert.True(preserved.Ok, string.Join("\n", preserved.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(preserved.File.ToByteArray());
        AssertTwoCellPicture(preserved.File.ToByteArray(), "Updated two-cell mark", "Updated two-cell performance", 4, 3, 9, 8, Xdr.EditAsValues.Absolute, residual: true);

        var withoutEditAs = Import(preserved.File.ToByteArray());
        withoutEditAs.Artifact.Workbook.Worksheets[0].Images[0].TwoCellAnchor.ClearEditAs();
        var omitted = Export(withoutEditAs.Artifact);
        Assert.True(omitted.Ok, string.Join("\n", omitted.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(omitted.File.ToByteArray());
        AssertTwoCellPicture(omitted.File.ToByteArray(), "Updated two-cell mark", "Updated two-cell performance", 4, 3, 9, 8, null, residual: true);
        var omittedImport = Import(omitted.File.ToByteArray());
        Assert.False(Assert.Single(omittedImport.Artifact.Workbook.Worksheets[0].Images).TwoCellAnchor.HasEditAs);

        var changedKind = Import(source);
        var changedImage = changedKind.Artifact.Workbook.Worksheets[0].Images[0];
        changedImage.Anchor = new SpreadsheetOneCellAnchorArtifact { Row = 3, Column = 2, WidthEmu = 1_143_000, HeightEmu = 762_000 };
        changedImage.TwoCellAnchor = null;
        var rejected = Export(changedKind.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_spreadsheet_image_edit", Assert.Single(rejected.Diagnostics).Code);

        var invalidGeometry = TwoCellPictureExportRequest();
        invalidGeometry.Artifact.Workbook.Worksheets[0].Images[0].TwoCellAnchor.To.Row = 2;
        rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(invalidGeometry.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_image", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void ProtocolAuthorsImportsAndSourcePreservesAbsoluteWorksheetPictures()
    {
        var request = AbsolutePictureExportRequest();
        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.True(authored.Ok, string.Join("\n", authored.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(authored.File.ToByteArray());
        AssertAbsolutePicture(authored.File.ToByteArray(), "Quarter mark", "Quarterly performance", -190_500, 285_750, 1_143_000, 762_000, residual: false);

        var source = AddPictureResidual(authored.File.ToByteArray());
        var imported = Import(source);
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var image = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Images);
        Assert.Null(image.Anchor);
        Assert.Null(image.TwoCellAnchor);
        Assert.NotNull(image.AbsoluteAnchor);
        Assert.Equal(-190_500, image.AbsoluteAnchor.XEmu);
        Assert.Equal(285_750, image.AbsoluteAnchor.YEmu);
        Assert.Equal(1_143_000, image.AbsoluteAnchor.WidthEmu);
        Assert.Equal(762_000, image.AbsoluteAnchor.HeightEmu);

        image.Name = "Updated absolute mark";
        image.AltText = "Updated absolute performance";
        image.AbsoluteAnchor.XEmu = 381_000;
        image.AbsoluteAnchor.YEmu = -95_250;
        image.AbsoluteAnchor.WidthEmu = 1_524_000;
        image.AbsoluteAnchor.HeightEmu = 952_500;
        var preserved = Export(imported.Artifact);
        Assert.True(preserved.Ok, string.Join("\n", preserved.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(preserved.File.ToByteArray());
        AssertAbsolutePicture(preserved.File.ToByteArray(), "Updated absolute mark", "Updated absolute performance", 381_000, -95_250, 1_524_000, 952_500, residual: true);
        var reimported = Import(preserved.File.ToByteArray());
        Assert.True(reimported.Ok, string.Join("\n", reimported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal(381_000, Assert.Single(reimported.Artifact.Workbook.Worksheets[0].Images).AbsoluteAnchor.XEmu);

        var changedKind = Import(source);
        var changedImage = changedKind.Artifact.Workbook.Worksheets[0].Images[0];
        changedImage.Anchor = new SpreadsheetOneCellAnchorArtifact { Row = 3, Column = 2, WidthEmu = 1_143_000, HeightEmu = 762_000 };
        changedImage.AbsoluteAnchor = null;
        var rejected = Export(changedKind.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_spreadsheet_image_edit", Assert.Single(rejected.Diagnostics).Code);

        var invalidExtent = AbsolutePictureExportRequest();
        invalidExtent.Artifact.Workbook.Worksheets[0].Images[0].AbsoluteAnchor.WidthEmu = 0;
        rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(invalidExtent.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_image", Assert.Single(rejected.Diagnostics).Code);

        var invalidPosition = AbsolutePictureExportRequest();
        invalidPosition.Artifact.Workbook.Worksheets[0].Images[0].AbsoluteAnchor.XEmu = 95_250_000_001;
        rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(invalidPosition.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_image", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void ProtocolAuthorsImportsEditsRemovesAndBoundsWorksheetPictureCrop()
    {
        var request = CropPictureExportRequest();
        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.True(authored.Ok, string.Join("\n", authored.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(authored.File.ToByteArray());
        AssertPictureCrop(authored.File.ToByteArray(), 10_000, -5_000, 15_000, 20_000, locked: false);

        var source = AddPictureLocksResidual(authored.File.ToByteArray());
        var imported = Import(source);
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var image = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Images);
        Assert.Equal(10_000, image.Crop.LeftThousandthPercent);
        Assert.Equal(-5_000, image.Crop.TopThousandthPercent);
        Assert.Equal(15_000, image.Crop.RightThousandthPercent);
        Assert.Equal(20_000, image.Crop.BottomThousandthPercent);

        image.Crop.LeftThousandthPercent = 12_000;
        image.Crop.TopThousandthPercent = 7_000;
        image.Crop.RightThousandthPercent = -3_000;
        image.Crop.BottomThousandthPercent = 18_000;
        var edited = Export(imported.Artifact);
        Assert.True(edited.Ok, string.Join("\n", edited.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(edited.File.ToByteArray());
        AssertPictureCrop(edited.File.ToByteArray(), 12_000, 7_000, -3_000, 18_000, locked: true);
        var reimported = Import(edited.File.ToByteArray());
        Assert.True(reimported.Ok, string.Join("\n", reimported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal(-3_000, Assert.Single(reimported.Artifact.Workbook.Worksheets[0].Images).Crop.RightThousandthPercent);

        var removed = Import(edited.File.ToByteArray());
        removed.Artifact.Workbook.Worksheets[0].Images[0].Crop = null;
        var withoutCrop = Export(removed.Artifact);
        Assert.True(withoutCrop.Ok, string.Join("\n", withoutCrop.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(withoutCrop.File.ToByteArray());
        AssertPictureCrop(withoutCrop.File.ToByteArray(), null, null, null, null, locked: true);
        var removedImport = Import(withoutCrop.File.ToByteArray());
        Assert.True(removedImport.Ok, string.Join("\n", removedImport.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Null(Assert.Single(removedImport.Artifact.Workbook.Worksheets[0].Images).Crop);

        var invalidPair = CropPictureExportRequest();
        invalidPair.Artifact.Workbook.Worksheets[0].Images[0].Crop.LeftThousandthPercent = 60_000;
        invalidPair.Artifact.Workbook.Worksheets[0].Images[0].Crop.RightThousandthPercent = 40_000;
        var rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(invalidPair.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_image", Assert.Single(rejected.Diagnostics).Code);

        var invalidBound = CropPictureExportRequest();
        invalidBound.Artifact.Workbook.Worksheets[0].Images[0].Crop.LeftThousandthPercent = 100_001;
        rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(invalidBound.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_image", Assert.Single(rejected.Diagnostics).Code);

        var opaque = Import(SetPictureCrop(authored.File.ToByteArray(), 100_001, 0, 0, 0));
        Assert.True(opaque.Ok, string.Join("\n", opaque.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var opaqueImage = Assert.Single(opaque.Artifact.Workbook.Worksheets[0].Images);
        Assert.Null(opaqueImage.Crop);
        opaqueImage.Crop = new SpreadsheetImageCropArtifact { LeftThousandthPercent = 1_000 };
        rejected = Export(opaque.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_spreadsheet_image_edit", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void ProtocolAuthorsImportsEditsRemovesAndBoundsWorksheetPictureEffects()
    {
        var request = EffectsPictureExportRequest();
        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.True(authored.Ok, string.Join("\n", authored.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(authored.File.ToByteArray());
        AssertPictureEffects(authored.File.ToByteArray(), 65_000, true, 15_000, -10_000, locked: false);

        var source = AddPictureLocksResidual(authored.File.ToByteArray());
        var imported = Import(source);
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var image = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Images);
        Assert.True(image.Effects.Grayscale);
        Assert.Equal(15_000, image.Effects.Luminance.BrightnessThousandthPercent);
        Assert.Equal(-10_000, image.Effects.Luminance.ContrastThousandthPercent);
        Assert.True(image.Effects.HasOpacityThousandthPercent);
        Assert.Equal(65_000U, image.Effects.OpacityThousandthPercent);

        image.Effects.Grayscale = false;
        image.Effects.Luminance.BrightnessThousandthPercent = -20_000;
        image.Effects.Luminance.ContrastThousandthPercent = 25_000;
        image.Effects.OpacityThousandthPercent = 0;
        var edited = Export(imported.Artifact);
        Assert.True(edited.Ok, string.Join("\n", edited.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(edited.File.ToByteArray());
        AssertPictureEffects(edited.File.ToByteArray(), 0, false, -20_000, 25_000, locked: true);

        var removed = Import(edited.File.ToByteArray());
        removed.Artifact.Workbook.Worksheets[0].Images[0].Effects = null;
        var withoutEffects = Export(removed.Artifact);
        Assert.True(withoutEffects.Ok, string.Join("\n", withoutEffects.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(withoutEffects.File.ToByteArray());
        AssertPictureEffects(withoutEffects.File.ToByteArray(), null, false, null, null, locked: true);
        Assert.Null(Assert.Single(Import(withoutEffects.File.ToByteArray()).Artifact.Workbook.Worksheets[0].Images).Effects);

        var invalidLuminance = EffectsPictureExportRequest();
        invalidLuminance.Artifact.Workbook.Worksheets[0].Images[0].Effects.Luminance.BrightnessThousandthPercent = 100_001;
        var rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(invalidLuminance.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_image", Assert.Single(rejected.Diagnostics).Code);

        var invalidOpacity = EffectsPictureExportRequest();
        invalidOpacity.Artifact.Workbook.Worksheets[0].Images[0].Effects.OpacityThousandthPercent = 100_001;
        rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(invalidOpacity.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_image", Assert.Single(rejected.Diagnostics).Code);

        var empty = PictureExportRequest();
        empty.Artifact.Workbook.Worksheets[0].Images[0].Effects = new SpreadsheetImageEffectsArtifact();
        rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(empty.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_image", Assert.Single(rejected.Diagnostics).Code);

        var opaque = Import(AddDuplicatePictureEffect(authored.File.ToByteArray()));
        Assert.True(opaque.Ok, string.Join("\n", opaque.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var opaqueImage = Assert.Single(opaque.Artifact.Workbook.Worksheets[0].Images);
        Assert.Null(opaqueImage.Effects);
        opaqueImage.Name = "Opaque effect retained";
        var metadataOnly = Export(opaque.Artifact);
        Assert.True(metadataOnly.Ok, string.Join("\n", metadataOnly.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertPictureGrayscaleCount(metadataOnly.File.ToByteArray(), 2);
        var opaqueEdit = Import(metadataOnly.File.ToByteArray());
        opaqueEdit.Artifact.Workbook.Worksheets[0].Images[0].Effects = new SpreadsheetImageEffectsArtifact { Grayscale = true };
        rejected = Export(opaqueEdit.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_spreadsheet_image_edit", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void ProtocolAuthorsImportsEditsRemovesAndBoundsWorksheetPictureTransform()
    {
        var request = TransformPictureExportRequest();
        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
        Assert.True(authored.Ok, string.Join("\n", authored.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(authored.File.ToByteArray());
        AssertPictureTransform(authored.File.ToByteArray(), 1_830_000, true, false, residual: false);

        var source = AddPictureTransformResidual(AddPictureLocksResidual(authored.File.ToByteArray()));
        var imported = Import(source);
        Assert.True(imported.Ok, string.Join("\n", imported.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var image = Assert.Single(imported.Artifact.Workbook.Worksheets[0].Images);
        Assert.Equal(1_830_000, image.Transform.RotationAngle60000);
        Assert.True(image.Transform.FlipHorizontal);
        Assert.False(image.Transform.FlipVertical);
        Assert.True(image.Transform.HasFlipHorizontal);
        Assert.True(image.Transform.HasFlipVertical);

        image.Transform.RotationAngle60000 = -2_700_000;
        image.Transform.FlipHorizontal = false;
        image.Transform.FlipVertical = true;
        var edited = Export(imported.Artifact);
        Assert.True(edited.Ok, string.Join("\n", edited.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(edited.File.ToByteArray());
        AssertPictureTransform(edited.File.ToByteArray(), -2_700_000, false, true, residual: true);

        var removed = Import(edited.File.ToByteArray());
        removed.Artifact.Workbook.Worksheets[0].Images[0].Transform = null;
        var withoutTransform = Export(removed.Artifact);
        Assert.True(withoutTransform.Ok, string.Join("\n", withoutTransform.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertOffice2021Valid(withoutTransform.File.ToByteArray());
        AssertPictureTransform(withoutTransform.File.ToByteArray(), null, null, null, residual: true);
        Assert.Null(Assert.Single(Import(withoutTransform.File.ToByteArray()).Artifact.Workbook.Worksheets[0].Images).Transform);

        var invalidRotation = TransformPictureExportRequest();
        invalidRotation.Artifact.Workbook.Worksheets[0].Images[0].Transform.RotationAngle60000 = 21_600_001;
        var rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(invalidRotation.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_image", Assert.Single(rejected.Diagnostics).Code);

        var empty = PictureExportRequest();
        empty.Artifact.Workbook.Worksheets[0].Images[0].Transform = new SpreadsheetImageTransformArtifact();
        rejected = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(empty.ToByteArray()));
        Assert.False(rejected.Ok);
        Assert.Equal("invalid_spreadsheet_image", Assert.Single(rejected.Diagnostics).Code);

        var opaque = Import(SetPictureTransformRotation(authored.File.ToByteArray(), 21_600_001));
        Assert.True(opaque.Ok, string.Join("\n", opaque.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var opaqueImage = Assert.Single(opaque.Artifact.Workbook.Worksheets[0].Images);
        Assert.Null(opaqueImage.Transform);
        opaqueImage.Name = "Opaque transform retained";
        var metadataOnly = Export(opaque.Artifact);
        Assert.True(metadataOnly.Ok, string.Join("\n", metadataOnly.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        AssertPictureTransform(metadataOnly.File.ToByteArray(), 21_600_001, true, false, residual: false);
        var opaqueEdit = Import(metadataOnly.File.ToByteArray());
        opaqueEdit.Artifact.Workbook.Worksheets[0].Images[0].Transform = new SpreadsheetImageTransformArtifact { RotationAngle60000 = 600_000 };
        rejected = Export(opaqueEdit.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_spreadsheet_image_edit", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void SourceBoundWorksheetPicturesRejectSharedOrCrossFormatReplacement()
    {
        var authored = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(PictureExportRequest().ToByteArray()));
        Assert.True(authored.Ok, string.Join("\n", authored.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));

        var crossFormat = Import(authored.File.ToByteArray());
        Assert.True(crossFormat.Ok, string.Join("\n", crossFormat.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        var jpeg = WorksheetImageAsset([0xff, 0xd8, 0xff, 0xd9], "image/jpeg", "replacement.jpg");
        crossFormat.Artifact.Assets.Add(jpeg);
        crossFormat.Artifact.Workbook.Worksheets[0].Images[0].AssetId = jpeg.Id;
        var rejected = Export(crossFormat.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_spreadsheet_image_edit", Assert.Single(rejected.Diagnostics).Code);
        Assert.Contains("content type", Assert.Single(rejected.Diagnostics).Message, StringComparison.OrdinalIgnoreCase);

        var shared = Import(AddSharedPictureReference(authored.File.ToByteArray()));
        Assert.True(shared.Ok, string.Join("\n", shared.Diagnostics.Select(item => $"{item.Code}: {item.Message}")));
        Assert.Equal(2, shared.Artifact.Workbook.Worksheets[0].Images.Count);
        var png = WorksheetImageAsset(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAQAAABFaP0WAAAADUlEQVR42mNk+M/wHwAF/gL+3c5GAAAAAElFTkSuQmCC"), "image/png", "replacement.png");
        shared.Artifact.Assets.Add(png);
        shared.Artifact.Workbook.Worksheets[0].Images[0].AssetId = png.Id;
        rejected = Export(shared.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_spreadsheet_image_edit", Assert.Single(rejected.Diagnostics).Code);
        Assert.Contains("referenced by 2 picture blips", Assert.Single(rejected.Diagnostics).Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void WorksheetPicturesRejectInvalidAssetsAndGeometry()
    {
        var invalidAsset = PictureExportRequest();
        invalidAsset.Artifact.Assets[0].Data = ByteString.CopyFromUtf8("not a png");
        var response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(invalidAsset.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_spreadsheet_image_asset", Assert.Single(response.Diagnostics).Code);

        var invalidAnchor = PictureExportRequest();
        invalidAnchor.Artifact.Workbook.Worksheets[0].Images[0].Anchor.WidthEmu = 0;
        response = CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(invalidAnchor.ToByteArray()));
        Assert.False(response.Ok);
        Assert.Equal("invalid_spreadsheet_image", Assert.Single(response.Diagnostics).Code);
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

    private static CodecRequest PictureExportRequest()
    {
        var request = ExportRequest();
        var bytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=");
        var asset = WorksheetImageAsset(bytes, "image/png", "quarter-mark.png");
        request.Artifact.Assets.Add(asset);
        request.Artifact.Workbook.Worksheets[0].Images.Add(new SpreadsheetImageArtifact
        {
            Id = "worksheet/summary/image/quarter-mark",
            Name = "Quarter mark",
            AltText = "Quarterly performance",
            AssetId = asset.Id,
            Anchor = new SpreadsheetOneCellAnchorArtifact
            {
                Row = 3,
                Column = 2,
                RowOffsetEmu = 95_250,
                ColumnOffsetEmu = 47_625,
                WidthEmu = 1_143_000,
                HeightEmu = 762_000,
            },
        });
        return request;
    }

    private static CodecRequest ChartExportRequest()
    {
        var request = PictureExportRequest();
        var chart = new SpreadsheetChartArtifact
        {
            Id = "worksheet/summary/chart/quarter-trend",
            Name = "Quarter chart",
            Title = "Quarter trend",
            TitleTextStyle = new SpreadsheetChartTextStyleArtifact { FontSizePoints = 12 },
            Type = SpreadsheetChartType.Line,
            HasLegend = true,
            AbsoluteAnchor = new SpreadsheetAbsoluteAnchorArtifact { XEmu = 3_619_500, YEmu = 190_500, WidthEmu = 3_429_000, HeightEmu = 2_095_500 },
            XAxis = new SpreadsheetChartAxisArtifact { Title = "Quarter", NumberFormatCode = "@", TickLabelInterval = 2, TextStyle = new SpreadsheetChartTextStyleArtifact { FontSizePoints = 10 } },
            YAxis = new SpreadsheetChartAxisArtifact { Title = "Revenue", NumberFormatCode = "$#,##0.0", Minimum = 0, Maximum = 100, MajorUnit = 25, TextStyle = new SpreadsheetChartTextStyleArtifact { FontSizePoints = 9 } },
        };
        chart.Categories.Add(["Q1", "Q2"]);
        chart.Series.Add(new SpreadsheetChartSeriesArtifact
        {
            Name = "Revenue",
            CategoryFormula = "'Summary'!$A$1:$A$2",
            ValueFormula = "'Summary'!$B$1:$B$2",
            Fill = new SpreadsheetColor { Rgb = "F472B6" },
            Line = new SpreadsheetChartLineStyleArtifact
            {
                Color = new SpreadsheetColor { Rgb = "0EA5E9" },
                DashStyle = SpreadsheetChartLineDashStyle.Dashed,
                WidthPoints = 2,
            },
            Marker = new SpreadsheetChartMarkerArtifact { Symbol = SpreadsheetChartMarkerSymbol.Diamond, Size = 8 },
            Values = { 42.5, 85 },
        });
        request.Artifact.Workbook.Worksheets[0].Charts.Add(chart);
        return request;
    }

    private static CodecRequest TwoCellPictureExportRequest()
    {
        var request = PictureExportRequest();
        var image = request.Artifact.Workbook.Worksheets[0].Images[0];
        image.Anchor = null;
        image.TwoCellAnchor = new SpreadsheetTwoCellAnchorArtifact
        {
            From = new SpreadsheetCellMarkerArtifact { Row = 3, Column = 2, RowOffsetEmu = 95_250, ColumnOffsetEmu = 47_625 },
            To = new SpreadsheetCellMarkerArtifact { Row = 7, Column = 6, RowOffsetEmu = 190_500, ColumnOffsetEmu = 142_875 },
            EditAs = SpreadsheetTwoCellEditAs.OneCell,
        };
        return request;
    }

    private static CodecRequest AbsolutePictureExportRequest()
    {
        var request = PictureExportRequest();
        var image = request.Artifact.Workbook.Worksheets[0].Images[0];
        image.Anchor = null;
        image.AbsoluteAnchor = new SpreadsheetAbsoluteAnchorArtifact
        {
            XEmu = -190_500,
            YEmu = 285_750,
            WidthEmu = 1_143_000,
            HeightEmu = 762_000,
        };
        return request;
    }

    private static CodecRequest CropPictureExportRequest()
    {
        var request = PictureExportRequest();
        request.Artifact.Workbook.Worksheets[0].Images[0].Crop = new SpreadsheetImageCropArtifact
        {
            LeftThousandthPercent = 10_000,
            TopThousandthPercent = -5_000,
            RightThousandthPercent = 15_000,
            BottomThousandthPercent = 20_000,
        };
        return request;
    }

    private static CodecRequest EffectsPictureExportRequest()
    {
        var request = PictureExportRequest();
        request.Artifact.Workbook.Worksheets[0].Images[0].Effects = new SpreadsheetImageEffectsArtifact
        {
            Grayscale = true,
            Luminance = new SpreadsheetImageLuminanceEffectArtifact
            {
                BrightnessThousandthPercent = 15_000,
                ContrastThousandthPercent = -10_000,
            },
            OpacityThousandthPercent = 65_000,
        };
        return request;
    }

    private static CodecRequest TransformPictureExportRequest()
    {
        var request = PictureExportRequest();
        request.Artifact.Workbook.Worksheets[0].Images[0].Transform = new SpreadsheetImageTransformArtifact
        {
            RotationAngle60000 = 1_830_000,
            FlipHorizontal = true,
            FlipVertical = false,
        };
        return request;
    }

    private static Asset WorksheetImageAsset(byte[] bytes, string contentType, string fileName)
    {
        var digest = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        return new Asset
        {
            Id = $"asset/workbook/image/{digest}",
            FileName = fileName,
            ContentType = contentType,
            Data = ByteString.CopyFrom(bytes),
            Sha256 = digest,
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

    private static CodecRequest DefinedNameExportRequest()
    {
        var request = ExportRequest();
        request.Artifact.Workbook.DefinedNames.Add(new SpreadsheetDefinedNameArtifact
        {
            Id = "defined-name/global-revenue",
            Name = "RevenueData",
            RefersTo = "Summary!$B$1:$B$2",
            Comment = "Revenue body",
            Hidden = false,
        });
        request.Artifact.Workbook.DefinedNames.Add(new SpreadsheetDefinedNameArtifact
        {
            Id = "defined-name/local-revenue",
            Name = "RevenueData",
            RefersTo = "Summary!$B$1",
            ScopeSheetName = "Summary",
            Hidden = true,
        });
        return request;
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
        var sort = new SpreadsheetTableSortStateArtifact { Reference = "A2:B3", CaseSensitive = true, SortMethod = "stroke" };
        sort.Conditions.Add(new SpreadsheetTableSortConditionArtifact { Reference = "B2:B3", Descending = true, CustomList = "High,Medium,Low" });
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

    private static byte[] AddChartResidual(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var chartPart = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ChartParts.Single();
            var chart = XDocument.Parse(ReadPartText(chartPart));
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            XNamespace fixture = "urn:openchestnut:worksheet-chart";
            chart.Root!.Add(new XElement(c + "extLst", new XElement(c + "ext", new XAttribute("uri", "{B6E80C28-38CE-4C9A-91B6-784D65016615}"), new XElement(fixture + "probe", "preserve-me"))));
            using var output = chartPart.GetStream(FileMode.Create, FileAccess.Write);
            chart.Save(output, SaveOptions.DisableFormatting);
        }
        return stream.ToArray();
    }

    private static byte[] SetChartSeriesNameReference(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var chartPart = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ChartParts.Single();
            var chart = XDocument.Parse(ReadPartText(chartPart));
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            var tx = chart.Descendants(c + "ser").Single().Element(c + "tx")!;
            tx.ReplaceNodes(new XElement(c + "strRef",
                new XElement(c + "f", "'Summary'!$B$1"),
                new XElement(c + "strCache",
                    new XElement(c + "ptCount", new XAttribute("val", 1)),
                    new XElement(c + "pt", new XAttribute("idx", 0), new XElement(c + "v", "Revenue from cache")))));
            using var output = chartPart.GetStream(FileMode.Create, FileAccess.Write);
            chart.Save(output, SaveOptions.DisableFormatting);
        }
        return stream.ToArray();
    }

    private static byte[] SetChartValueAxisLogarithmic(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var chartPart = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ChartParts.Single();
            var chart = XDocument.Parse(ReadPartText(chartPart));
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            chart.Descendants(c + "valAx").Single().Element(c + "scaling")!.AddFirst(new XElement(c + "logBase", new XAttribute("val", 10)));
            using var output = chartPart.GetStream(FileMode.Create, FileAccess.Write);
            chart.Save(output, SaveOptions.DisableFormatting);
        }
        return stream.ToArray();
    }

    private static byte[] SetChartSeriesSchemeFill(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var chartPart = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ChartParts.Single();
            var chart = XDocument.Parse(ReadPartText(chartPart));
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            chart.Descendants(c + "ser").Single().Element(c + "spPr")!.Element(a + "solidFill")!
                .ReplaceNodes(new XElement(a + "schemeClr", new XAttribute("val", "accent1")));
            using var output = chartPart.GetStream(FileMode.Create, FileAccess.Write);
            chart.Save(output, SaveOptions.DisableFormatting);
        }
        return stream.ToArray();
    }

    private static byte[] SetChartSeriesComplexLine(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var chartPart = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ChartParts.Single();
            var chart = XDocument.Parse(ReadPartText(chartPart));
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            chart.Descendants(c + "ser").Single().Element(c + "spPr")!.Element(a + "ln")!
                .Add(new XElement(a + "headEnd", new XAttribute("type", "triangle")));
            using var output = chartPart.GetStream(FileMode.Create, FileAccess.Write);
            chart.Save(output, SaveOptions.DisableFormatting);
        }
        return stream.ToArray();
    }

    private static byte[] SetChartSeriesComplexMarker(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var chartPart = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ChartParts.Single();
            var chart = XDocument.Parse(ReadPartText(chartPart));
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            chart.Descendants(c + "ser").Single().Element(c + "marker")!
                .Add(new XElement(c + "spPr", new XElement(a + "solidFill", new XElement(a + "schemeClr", new XAttribute("val", "accent1")))));
            using var output = chartPart.GetStream(FileMode.Create, FileAccess.Write);
            chart.Save(output, SaveOptions.DisableFormatting);
        }
        return stream.ToArray();
    }

    private static byte[] SetChartAxisComplexTextStyle(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var chartPart = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ChartParts.Single();
            var chart = XDocument.Parse(ReadPartText(chartPart));
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            chart.Descendants(c + "catAx").Single().Element(c + "txPr")!.Descendants(a + "defRPr").Single()
                .Add(new XElement(a + "solidFill", new XElement(a + "schemeClr", new XAttribute("val", "accent1"))));
            using var output = chartPart.GetStream(FileMode.Create, FileAccess.Write);
            chart.Save(output, SaveOptions.DisableFormatting);
        }
        return stream.ToArray();
    }

    private static void AssertChartAxes(byte[] bytes, string xTitle, string xFormat, uint xInterval, string yTitle, string yFormat, double yMinimum, double yMaximum, double yMajorUnit)
    {
        var chart = XDocument.Parse(ReadChartXml(bytes));
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var xAxis = chart.Descendants(c + "catAx").Single();
        var yAxis = chart.Descendants(c + "valAx").Single();
        Assert.Equal(xTitle, string.Concat(xAxis.Element(c + "title")!.Descendants(a + "t").Select(item => item.Value)));
        Assert.Equal(xFormat, (string?)xAxis.Element(c + "numFmt")!.Attribute("formatCode"));
        Assert.Equal(xInterval.ToString(CultureInfo.InvariantCulture), (string?)xAxis.Element(c + "tickLblSkip")!.Attribute("val"));
        Assert.Equal(yTitle, string.Concat(yAxis.Element(c + "title")!.Descendants(a + "t").Select(item => item.Value)));
        Assert.Equal(yFormat, (string?)yAxis.Element(c + "numFmt")!.Attribute("formatCode"));
        Assert.Equal(yMinimum.ToString("R", CultureInfo.InvariantCulture), (string?)yAxis.Element(c + "scaling")!.Element(c + "min")!.Attribute("val"));
        Assert.Equal(yMaximum.ToString("R", CultureInfo.InvariantCulture), (string?)yAxis.Element(c + "scaling")!.Element(c + "max")!.Attribute("val"));
        Assert.Equal(yMajorUnit.ToString("R", CultureInfo.InvariantCulture), (string?)yAxis.Element(c + "majorUnit")!.Attribute("val"));
    }

    private static void AssertChartTextStyles(byte[] bytes, double? titleFontSize, double? xFontSize, double? yFontSize)
    {
        var chart = XDocument.Parse(ReadChartXml(bytes));
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        static double? FontSize(XElement? properties) => uint.TryParse((string?)properties?.Attribute("sz"), NumberStyles.None, CultureInfo.InvariantCulture, out var size) ? size / 100d : null;
        Assert.Equal(titleFontSize, FontSize(chart.Root!.Element(c + "chart")!.Element(c + "title")?.Descendants(a + "rPr").SingleOrDefault()));
        Assert.Equal(xFontSize, FontSize(chart.Descendants(c + "catAx").Single().Element(c + "txPr")?.Descendants(a + "defRPr").SingleOrDefault()));
        Assert.Equal(yFontSize, FontSize(chart.Descendants(c + "valAx").Single().Element(c + "txPr")?.Descendants(a + "defRPr").SingleOrDefault()));
    }

    private static void AssertChartSeriesFill(byte[] bytes, string? expectedRgb)
    {
        var chart = XDocument.Parse(ReadChartXml(bytes));
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var solidFill = chart.Descendants(c + "ser").Single().Element(c + "spPr")?.Element(a + "solidFill");
        if (expectedRgb is null) Assert.Null(solidFill);
        else Assert.Equal(expectedRgb, (string?)solidFill!.Element(a + "srgbClr")!.Attribute("val"));
    }

    private static void AssertChartSeriesLine(byte[] bytes, string? expectedRgb, string? expectedDash, double? expectedWidthPoints)
    {
        var chart = XDocument.Parse(ReadChartXml(bytes));
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var line = chart.Descendants(c + "ser").Single().Element(c + "spPr")?.Element(a + "ln");
        if (expectedRgb is null && expectedDash is null && expectedWidthPoints is null) { Assert.Null(line); return; }
        Assert.NotNull(line);
        Assert.Equal(expectedRgb, (string?)line!.Element(a + "solidFill")?.Element(a + "srgbClr")?.Attribute("val"));
        Assert.Equal(expectedDash, (string?)line.Element(a + "prstDash")?.Attribute("val"));
        Assert.Equal(expectedWidthPoints is null ? null : Math.Round(expectedWidthPoints.Value * 12_700, MidpointRounding.AwayFromZero).ToString(CultureInfo.InvariantCulture), (string?)line.Attribute("w"));
    }

    private static void AssertChartSeriesMarker(byte[] bytes, string? expectedSymbol, uint? expectedSize)
    {
        var chart = XDocument.Parse(ReadChartXml(bytes));
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        var marker = chart.Descendants(c + "ser").Single().Element(c + "marker");
        if (expectedSymbol is null && expectedSize is null) { Assert.Null(marker); return; }
        Assert.NotNull(marker);
        Assert.Equal(expectedSymbol, (string?)marker!.Element(c + "symbol")?.Attribute("val"));
        Assert.Equal(expectedSize?.ToString(CultureInfo.InvariantCulture), (string?)marker.Element(c + "size")?.Attribute("val"));
    }

    private static string ReadChartXml(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        return ReadPartText(document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ChartParts.Single());
    }

    private static string ReadPartText(OpenXmlPart part)
    {
        using var reader = new StreamReader(part.GetStream(FileMode.Open, FileAccess.Read));
        return reader.ReadToEnd();
    }

    private static byte[] AddPictureResidual(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var drawing = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing!;
            var picture = drawing.Descendants<Xdr.Picture>().Single();
            picture.NonVisualPictureProperties!.NonVisualPictureDrawingProperties!.Append(new A.PictureLocks { NoChangeAspect = true });
            var fill = picture.BlipFill!;
            fill.InsertAfter(new A.SourceRectangle { Left = 1_000, Top = 2_000 }, fill.Blip!);
            drawing.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddPictureLocksResidual(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var drawing = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing!;
            var picture = drawing.Descendants<Xdr.Picture>().Single();
            picture.NonVisualPictureProperties!.NonVisualPictureDrawingProperties!.Append(new A.PictureLocks { NoChangeAspect = true });
            var extension = new A.BlipExtension { Uri = "{E8D2B7F2-4A0E-4F79-9E08-9BFC0B4C25C6}" };
            extension.Append(document.WorkbookPart.WorksheetParts.Single().DrawingsPart!.CreateUnknownElement("<fixture:probe xmlns:fixture=\"urn:openchestnut:picture-effect\">preserve-me</fixture:probe>"));
            picture.BlipFill!.Blip!.Append(new A.BlipExtensionList(extension));
            drawing.Save();
        }
        return stream.ToArray();
    }

    private static byte[] SetPictureCrop(byte[] bytes, int left, int top, int right, int bottom)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var drawing = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing!;
            var fill = drawing.Descendants<Xdr.Picture>().Single().BlipFill!;
            var crop = fill.GetFirstChild<A.SourceRectangle>() ?? throw new InvalidOperationException("Expected an authored source rectangle.");
            crop.Left = left;
            crop.Top = top;
            crop.Right = right;
            crop.Bottom = bottom;
            drawing.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddDuplicatePictureEffect(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var drawing = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing!;
            drawing.Descendants<A.Blip>().Single().Append(new A.Grayscale());
            drawing.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddPictureTransformResidual(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var drawing = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing!;
            var transform = drawing.Descendants<Xdr.Picture>().Single().ShapeProperties!.GetFirstChild<A.Transform2D>()!;
            transform.Append(new A.Offset { X = 123_456, Y = 234_567 });
            transform.Append(new A.Extents { Cx = 1_111_111, Cy = 2_222_222 });
            drawing.Save();
        }
        return stream.ToArray();
    }

    private static byte[] SetPictureTransformRotation(byte[] bytes, int rotation)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var drawing = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing!;
            drawing.Descendants<Xdr.Picture>().Single().ShapeProperties!.GetFirstChild<A.Transform2D>()!.Rotation = rotation;
            drawing.Save();
        }
        return stream.ToArray();
    }

    private static byte[] AddSharedPictureReference(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var drawing = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing!;
            var clone = (Xdr.OneCellAnchor)drawing.Elements<Xdr.OneCellAnchor>().Single().CloneNode(true);
            var nonVisual = clone.Descendants<Xdr.NonVisualDrawingProperties>().Single();
            nonVisual.Id = 99U;
            nonVisual.Name = "Shared quarter mark";
            drawing.Append(clone);
            drawing.Save();
        }
        return stream.ToArray();
    }

    private static byte[] ReadPictureBytes(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = SpreadsheetDocument.Open(stream, false);
        using var image = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ImageParts.Single().GetStream(FileMode.Open, FileAccess.Read);
        using var copy = new MemoryStream();
        image.CopyTo(copy);
        return copy.ToArray();
    }

    private static string ReadPicturePartPath(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = SpreadsheetDocument.Open(stream, false);
        return document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ImageParts.Single().Uri.OriginalString;
    }

    private static void AssertPicture(byte[] bytes, string name, string description, uint row, uint column, long width, long height, bool residual)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = document.WorkbookPart!.WorksheetParts.Single();
        Assert.NotNull(worksheetPart.Worksheet!.GetFirstChild<Drawing>());
        var drawing = worksheetPart.DrawingsPart!.WorksheetDrawing!;
        var anchor = Assert.Single(drawing.Elements<Xdr.OneCellAnchor>());
        Assert.Equal(row.ToString(System.Globalization.CultureInfo.InvariantCulture), anchor.FromMarker!.RowId!.Text);
        Assert.Equal(column.ToString(System.Globalization.CultureInfo.InvariantCulture), anchor.FromMarker.ColumnId!.Text);
        Assert.Equal(width, anchor.Extent!.Cx!.Value);
        Assert.Equal(height, anchor.Extent.Cy!.Value);
        var picture = anchor.GetFirstChild<Xdr.Picture>()!;
        Assert.Equal(name, picture.NonVisualPictureProperties!.NonVisualDrawingProperties!.Name!.Value);
        Assert.Equal(description, picture.NonVisualPictureProperties.NonVisualDrawingProperties.Description!.Value);
        Assert.Single(worksheetPart.DrawingsPart.ImageParts);
        var locks = picture.NonVisualPictureProperties.NonVisualPictureDrawingProperties!.GetFirstChild<A.PictureLocks>();
        var crop = picture.BlipFill!.GetFirstChild<A.SourceRectangle>();
        if (residual)
        {
            Assert.True(locks!.NoChangeAspect!.Value);
            Assert.Equal(1_000, crop!.Left!.Value);
            Assert.Equal(2_000, crop.Top!.Value);
        }
        else
        {
            Assert.Null(locks);
            Assert.Null(crop);
        }
    }

    private static void AssertTwoCellPicture(byte[] bytes, string name, string description, uint fromRow, uint fromColumn, uint toRow, uint toColumn, Xdr.EditAsValues? editAs, bool residual)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = document.WorkbookPart!.WorksheetParts.Single();
        Assert.NotNull(worksheetPart.Worksheet!.GetFirstChild<Drawing>());
        var anchor = Assert.Single(worksheetPart.DrawingsPart!.WorksheetDrawing!.Elements<Xdr.TwoCellAnchor>());
        Assert.Equal(fromRow.ToString(System.Globalization.CultureInfo.InvariantCulture), anchor.FromMarker!.RowId!.Text);
        Assert.Equal(fromColumn.ToString(System.Globalization.CultureInfo.InvariantCulture), anchor.FromMarker.ColumnId!.Text);
        Assert.Equal(toRow.ToString(System.Globalization.CultureInfo.InvariantCulture), anchor.ToMarker!.RowId!.Text);
        Assert.Equal(toColumn.ToString(System.Globalization.CultureInfo.InvariantCulture), anchor.ToMarker.ColumnId!.Text);
        Assert.Equal(editAs, anchor.EditAs?.Value);
        var picture = anchor.GetFirstChild<Xdr.Picture>()!;
        Assert.Equal(name, picture.NonVisualPictureProperties!.NonVisualDrawingProperties!.Name!.Value);
        Assert.Equal(description, picture.NonVisualPictureProperties.NonVisualDrawingProperties.Description!.Value);
        Assert.Single(worksheetPart.DrawingsPart.ImageParts);
        var locks = picture.NonVisualPictureProperties.NonVisualPictureDrawingProperties!.GetFirstChild<A.PictureLocks>();
        var crop = picture.BlipFill!.GetFirstChild<A.SourceRectangle>();
        if (residual)
        {
            Assert.True(locks!.NoChangeAspect!.Value);
            Assert.Equal(1_000, crop!.Left!.Value);
            Assert.Equal(2_000, crop.Top!.Value);
        }
        else
        {
            Assert.Null(locks);
            Assert.Null(crop);
        }
    }

    private static void AssertAbsolutePicture(byte[] bytes, string name, string description, long x, long y, long width, long height, bool residual)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = document.WorkbookPart!.WorksheetParts.Single();
        Assert.NotNull(worksheetPart.Worksheet!.GetFirstChild<Drawing>());
        var anchor = Assert.Single(worksheetPart.DrawingsPart!.WorksheetDrawing!.Elements<Xdr.AbsoluteAnchor>());
        Assert.Equal(x, anchor.Position!.X!.Value);
        Assert.Equal(y, anchor.Position.Y!.Value);
        Assert.Equal(width, anchor.Extent!.Cx!.Value);
        Assert.Equal(height, anchor.Extent.Cy!.Value);
        var picture = anchor.GetFirstChild<Xdr.Picture>()!;
        Assert.Equal(name, picture.NonVisualPictureProperties!.NonVisualDrawingProperties!.Name!.Value);
        Assert.Equal(description, picture.NonVisualPictureProperties.NonVisualDrawingProperties.Description!.Value);
        Assert.Single(worksheetPart.DrawingsPart.ImageParts);
        var locks = picture.NonVisualPictureProperties.NonVisualPictureDrawingProperties!.GetFirstChild<A.PictureLocks>();
        var crop = picture.BlipFill!.GetFirstChild<A.SourceRectangle>();
        if (residual)
        {
            Assert.True(locks!.NoChangeAspect!.Value);
            Assert.Equal(1_000, crop!.Left!.Value);
            Assert.Equal(2_000, crop.Top!.Value);
        }
        else
        {
            Assert.Null(locks);
            Assert.Null(crop);
        }
    }

    private static void AssertPictureCrop(byte[] bytes, int? left, int? top, int? right, int? bottom, bool locked)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = SpreadsheetDocument.Open(stream, false);
        var picture = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing!.Descendants<Xdr.Picture>().Single();
        var blip = picture.BlipFill!.Blip!;
        var crop = picture.BlipFill!.GetFirstChild<A.SourceRectangle>();
        if (left.HasValue)
        {
            Assert.NotNull(crop);
            Assert.Equal(left.Value, crop!.Left!.Value);
            Assert.Equal(top!.Value, crop.Top!.Value);
            Assert.Equal(right!.Value, crop.Right!.Value);
            Assert.Equal(bottom!.Value, crop.Bottom!.Value);
            var children = picture.BlipFill.ChildElements.ToList();
            Assert.True(children.IndexOf(crop) > children.IndexOf(picture.BlipFill.Blip!));
            Assert.True(children.IndexOf(crop) < children.IndexOf(picture.BlipFill.GetFirstChild<A.Stretch>()!));
        }
        else
        {
            Assert.Null(crop);
        }
        var locks = picture.NonVisualPictureProperties!.NonVisualPictureDrawingProperties!.GetFirstChild<A.PictureLocks>();
        if (locked)
        {
            Assert.True(locks!.NoChangeAspect!.Value);
            Assert.Contains("preserve-me", blip.GetFirstChild<A.BlipExtensionList>()!.OuterXml);
        }
        else
        {
            Assert.Null(locks);
            Assert.Null(blip.GetFirstChild<A.BlipExtensionList>());
        }
    }

    private static void AssertPictureEffects(byte[] bytes, int? opacity, bool grayscale, int? brightness, int? contrast, bool locked)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = SpreadsheetDocument.Open(stream, false);
        var picture = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing!.Descendants<Xdr.Picture>().Single();
        var blip = picture.BlipFill!.Blip!;
        var alpha = blip.GetFirstChild<A.AlphaModulationFixed>();
        var gray = blip.GetFirstChild<A.Grayscale>();
        var luminance = blip.GetFirstChild<A.LuminanceEffect>();
        if (opacity.HasValue) Assert.Equal(opacity.Value, alpha!.Amount!.Value);
        else Assert.Null(alpha);
        if (grayscale) Assert.NotNull(gray);
        else Assert.Null(gray);
        if (brightness.HasValue)
        {
            Assert.Equal(brightness.Value, luminance!.Brightness!.Value);
            Assert.Equal(contrast!.Value, luminance.Contrast!.Value);
        }
        else Assert.Null(luminance);
        var children = blip.ChildElements.ToList();
        if (alpha is not null && gray is not null) Assert.True(children.IndexOf(alpha) < children.IndexOf(gray));
        if (gray is not null && luminance is not null) Assert.True(children.IndexOf(gray) < children.IndexOf(luminance));
        if (alpha is not null && luminance is not null) Assert.True(children.IndexOf(alpha) < children.IndexOf(luminance));
        var locks = picture.NonVisualPictureProperties!.NonVisualPictureDrawingProperties!.GetFirstChild<A.PictureLocks>();
        if (locked)
        {
            Assert.True(locks!.NoChangeAspect!.Value);
            Assert.Contains("preserve-me", blip.GetFirstChild<A.BlipExtensionList>()!.OuterXml);
        }
        else
        {
            Assert.Null(locks);
            Assert.Null(blip.GetFirstChild<A.BlipExtensionList>());
        }
    }

    private static void AssertPictureGrayscaleCount(byte[] bytes, int expected)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = SpreadsheetDocument.Open(stream, false);
        var blip = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing!.Descendants<A.Blip>().Single();
        Assert.Equal(expected, blip.Elements<A.Grayscale>().Count());
    }

    private static void AssertPictureTransform(byte[] bytes, int? rotation, bool? flipHorizontal, bool? flipVertical, bool residual)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = SpreadsheetDocument.Open(stream, false);
        var picture = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing!.Descendants<Xdr.Picture>().Single();
        var transform = picture.ShapeProperties!.GetFirstChild<A.Transform2D>();
        Assert.NotNull(transform);
        if (rotation.HasValue) Assert.Equal(rotation.Value, transform!.Rotation!.Value);
        else Assert.Null(transform!.Rotation);
        if (flipHorizontal.HasValue) Assert.Equal(flipHorizontal.Value, transform.HorizontalFlip!.Value);
        else Assert.Null(transform.HorizontalFlip);
        if (flipVertical.HasValue) Assert.Equal(flipVertical.Value, transform.VerticalFlip!.Value);
        else Assert.Null(transform.VerticalFlip);
        var shapeChildren = picture.ShapeProperties.ChildElements.ToList();
        Assert.True(shapeChildren.IndexOf(transform) < shapeChildren.IndexOf(picture.ShapeProperties.GetFirstChild<A.PresetGeometry>()!));
        if (residual)
        {
            Assert.Equal(123_456L, transform.Offset!.X!.Value);
            Assert.Equal(234_567L, transform.Offset.Y!.Value);
            Assert.Equal(1_111_111L, transform.Extents!.Cx!.Value);
            Assert.Equal(2_222_222L, transform.Extents.Cy!.Value);
        }
        else
        {
            Assert.Null(transform.Offset);
            Assert.Null(transform.Extents);
        }
    }

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

    private static byte[] SetCellMetadataIndex(byte[] bytes, string reference, uint index)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            var cell = worksheet.Descendants<Cell>().Single(item => item.CellReference?.Value == reference);
            cell.CellMetaIndex = index;
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

    private static byte[] SetWorksheetSortConditionReferences(byte[] bytes, params string[] references)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            var conditions = worksheet.GetFirstChild<SortState>()!.Elements<SortCondition>().ToArray();
            Assert.Equal(conditions.Length, references.Length);
            for (var index = 0; index < conditions.Length; index++) conditions[index].Reference = references[index];
            worksheet.Save();
        }
        return stream.ToArray();
    }

    private static string ReadWorksheetSortXml(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        return document.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<SortState>()!.OuterXml;
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

    private static byte[] AddOpaqueDefinedName(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var workbook = document.WorkbookPart!.Workbook!;
            var names = workbook.DefinedNames ?? workbook.InsertAfter(new DefinedNames(), workbook.Sheets);
            names.Append(new DefinedName("42") { Name = "OpaqueConstant", Comment = "Preserve exactly" });
            workbook.Save();
        }
        return stream.ToArray();
    }

    private static string ReadDefinedNameXml(byte[] bytes, string name)
    {
        using var stream = new MemoryStream(bytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        return document.WorkbookPart!.Workbook!.DefinedNames!.Elements<DefinedName>()
            .Single(item => item.Name?.Value == name).OuterXml;
    }

    private static byte[] SetCalculationProfile(byte[] bytes, Action<CalculationProperties> mutate)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var workbook = document.WorkbookPart!.Workbook!;
            var calculation = workbook.CalculationProperties ?? new CalculationProperties();
            workbook.CalculationProperties = calculation;
            mutate(calculation);
            workbook.Save();
        }
        return stream.ToArray();
    }

    private static string ReadCalculationXml(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        return document.WorkbookPart!.Workbook!.CalculationProperties!.OuterXml;
    }

    private static byte[] SetRawSheetState(byte[] bytes, string sheetName, string state)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("xl/workbook.xml") ?? throw new InvalidOperationException("Workbook part is missing.");
            XDocument workbook;
            using (var reader = new StreamReader(entry.Open())) workbook = XDocument.Parse(reader.ReadToEnd(), LoadOptions.PreserveWhitespace);
            XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var sheet = workbook.Descendants(spreadsheet + "sheet").Single(item => (string?)item.Attribute("name") == sheetName);
            sheet.SetAttributeValue("state", state);
            entry.Delete();
            var replacement = archive.CreateEntry("xl/workbook.xml");
            using var writer = new StreamWriter(replacement.Open());
            writer.Write(workbook.ToString(SaveOptions.DisableFormatting));
        }
        return stream.ToArray();
    }

    private static byte[] MutateWorkbookViews(byte[] bytes, Action<BookViews> mutate)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var views = document.WorkbookPart!.Workbook!.BookViews ?? throw new InvalidOperationException("Workbook views are missing.");
            mutate(views);
            document.WorkbookPart.Workbook.Save();
        }
        return stream.ToArray();
    }

    private static WorkbookView[] ReadWorkbookViews(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        return document.WorkbookPart!.Workbook!.BookViews!.Elements<WorkbookView>().Select(view => (WorkbookView)view.CloneNode(true)).ToArray();
    }

    private static byte[] MutateWorksheetViews(byte[] bytes, int sheetOrdinal, Action<SheetViews> mutate)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var worksheet = WorksheetPartAt(document.WorkbookPart!, sheetOrdinal).Worksheet!;
            var views = worksheet.SheetViews ?? throw new InvalidOperationException("Worksheet views are missing.");
            mutate(views);
            worksheet.Save();
        }
        return stream.ToArray();
    }

    private static SheetView?[] ReadWorksheetViews(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart!;
        var count = workbookPart.Workbook!.Sheets!.Elements<Sheet>().Count();
        return Enumerable.Range(0, count)
            .Select(index => WorksheetPartAt(workbookPart, index).Worksheet!.SheetViews?.Elements<SheetView>().FirstOrDefault())
            .Select(view => view is null ? null : (SheetView)view.CloneNode(true))
            .ToArray();
    }

    private static SheetView[][] ReadWorksheetViewMatrix(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart!;
        var count = workbookPart.Workbook!.Sheets!.Elements<Sheet>().Count();
        return Enumerable.Range(0, count)
            .Select(index => WorksheetPartAt(workbookPart, index).Worksheet!.SheetViews!.Elements<SheetView>()
                .Select(view => (SheetView)view.CloneNode(true)).ToArray())
            .ToArray();
    }

    private static WorksheetPart WorksheetPartAt(WorkbookPart workbookPart, int ordinal)
    {
        var sheet = workbookPart.Workbook!.Sheets!.Elements<Sheet>().ElementAt(ordinal);
        return (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
    }

    private static byte[] AddQueryTableGraph(
        byte[] bytes,
        bool addUnsupportedRelationship = false,
        bool opaqueRefresh = false,
        bool opaqueDeletedFields = false,
        bool opaqueSort = false,
        bool opaqueConnection = false)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var workbookPart = document.WorkbookPart!;
            var connectionsPart = workbookPart.AddNewPart<ConnectionsPart>("rIdConnections");
            var connectionType = opaqueConnection ? 1 : 5;
            WritePart(connectionsPart, $$"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <x:connections xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:fixture="urn:open-office-artifact-tool:query-fixture">
                  <x:connection id="7" name="Fixture warehouse" description="Read-only warehouse source" type="{{connectionType}}" refreshedVersion="8" keepAlive="0" interval="30" background="1" refreshOnLoad="0" saveData="1" savePassword="0" credentials="integrated">
                    <x:dbPr connection="Provider=Fixture.Provider;Data Source=fixture.invalid" command="SELECT Region, Revenue FROM Sales" commandType="2"/>
                    <x:extLst><x:ext uri="{E5A74D42-D212-4CC7-9D5B-A7393F4D8A61}"><fixture:connectionOpaque value="kept"/></x:ext></x:extLst>
                  </x:connection>
                </x:connections>
                """);
            var tablePart = workbookPart.WorksheetParts.Single().TableDefinitionParts.Single();
            if (addUnsupportedRelationship)
                tablePart.AddExternalRelationship("urn:open-office-artifact-tool:unsupported-query-companion", new Uri("https://fixture.invalid/query"), "rIdUnsupportedQueryCompanion");
            var queryPart = tablePart.AddNewPart<QueryTablePart>("rIdQueryTable");
            var secondDeletedFieldName = opaqueDeletedFields ? "Legacy Region" : "Legacy Revenue";
            var sortOpaqueAttribute = opaqueSort ? " columnSort=\"1\"" : " sortMethod=\"stroke\"";
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
                      <x:sortCondition ref="B2:B3" descending="1" customList="North,South"/>
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
