using System.Globalization;
using System.Security.Cryptography;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

internal sealed record XlsxImportResult(ArtifactEnvelope Artifact, IReadOnlyList<Diagnostic> Diagnostics);
internal sealed record XlsxExportResult(byte[] File, IReadOnlyList<Diagnostic> Diagnostics);

internal static class XlsxCodec
{
    internal static XlsxExportResult Export(ArtifactEnvelope envelope, EffectiveCodecLimits limits)
    {
        if (envelope.ProtocolVersion != CodecProtocol.ProtocolVersion)
            throw new CodecException("unsupported_artifact_version", $"Artifact protocol version {envelope.ProtocolVersion} is unsupported.");
        if (envelope.Family != ArtifactFamily.Workbook || envelope.PayloadCase != ArtifactEnvelope.PayloadOneofCase.Workbook)
            throw new CodecException("invalid_workbook_artifact", "Artifact envelope does not contain a workbook payload.");
        var requiresSourcePreservation = RequiresSourcePreservation(envelope);
        if (requiresSourcePreservation && envelope.OpaqueOpc?.SourcePackage is not { Data.IsEmpty: false })
            throw new CodecException(
                "missing_source_package",
                "Source-bound XLSX export requires its validated source package snapshot.");
        ValidateWorkbookBudget(envelope.Workbook, limits);

        var opaqueCount = (envelope.OpaqueOpc?.Parts.Count ?? 0) + (envelope.OpaqueOpc?.PackageRelationships.Count ?? 0);
        if (requiresSourcePreservation)
            return ExportPreservingSource(envelope, limits, opaqueCount);

        var diagnostics = new List<Diagnostic>();

        using var stream = new MemoryStream();
        var imageAssets = new XlsxImageAssetCatalog(envelope.Assets, limits);
        using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, autoSave: true))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var dynamicArrays = new XlsxDynamicArrayCodec(workbookPart);
            if (envelope.Workbook.DateSystem == WorkbookDateSystem._1904)
                workbookPart.Workbook.WorkbookProperties = new WorkbookProperties { Date1904 = true };
            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            var theme = new XlsxThemeCodec(workbookPart);
            theme.Apply(envelope.Workbook.Theme, sourceBound: false);
            var styles = new XlsxCellStyleCodec(workbookPart);
            var worksheetFeatures = new XlsxWorksheetFeatureCodec(styles);
            var connections = new XlsxConnectionCodec(workbookPart);
            connections.Apply(envelope.Workbook.Connections, sourceBound: false);
            var calculation = new XlsxCalculationCodec(workbookPart);
            calculation.Apply(envelope.Workbook.Calculation, sourceBound: false);
            var sheetNames = envelope.Workbook.Worksheets.Select(sheet => sheet.Name).ToArray();
            var definedNames = new XlsxDefinedNameCodec(workbookPart, sheetNames);
            var drawings = new XlsxDrawingCodec(imageAssets);
            var charts = new XlsxChartCodec();
            var sparklines = new XlsxSparklineCodec();
            var nextTableId = 1U;
            var worksheetBindings = new List<(WorksheetPart Part, WorksheetArtifact Artifact)>();

            for (var index = 0; index < envelope.Workbook.Worksheets.Count; index++)
            {
                var source = envelope.Workbook.Worksheets[index];
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = BuildWorksheet(source, styles, dynamicArrays);
                new XlsxWorksheetProtectionCodec(worksheetPart).Apply(source.Protection, sourceBound: false);
                worksheetFeatures.ApplyRules(worksheetPart.Worksheet, source, sourceBound: false);
                var tables = new XlsxTableCodec(worksheetPart, workbookPart, styles, connections);
                tables.Apply(source.Tables, sourceBound: false, ref nextTableId);
                tables.Save();
                drawings.Apply(worksheetPart, source.Id, source.Images, sourceBound: false);
                charts.Apply(worksheetPart, source.Id, source.Charts, sourceBound: false);
                sparklines.Apply(worksheetPart, source.Id, source.SparklineGroups, sourceBound: false);
                worksheetPart.Worksheet.Save();
                sheets.Append(XlsxWorksheetMetadataCodec.Create(source, checked((uint)index), workbookPart.GetIdOfPart(worksheetPart)));
                worksheetBindings.Add((worksheetPart, source));
            }
            new XlsxPivotTableCodec(workbookPart).Apply(worksheetBindings, sourceBound: false);
            worksheetFeatures.ApplyThreadedComments(workbookPart, worksheetBindings, sourceBound: false);
            var workbookView = new XlsxWorkbookViewCodec(workbookPart, envelope.Workbook.Worksheets);
            workbookView.Apply(envelope.Workbook.View, envelope.Workbook.AdditionalViews, sourceBound: false, envelope.Workbook.Worksheets);
            definedNames.Apply(envelope.Workbook.DefinedNames, sourceBound: false, sheetNames);
            theme.Save();
            styles.Save();
            workbookPart.Workbook.Save();
        }

        var bytes = stream.ToArray();
        ValidateOutputBudget(bytes, limits);
        ValidateOffice2021(bytes);
        return new XlsxExportResult(bytes, diagnostics);
    }

    private static bool RequiresSourcePreservation(ArtifactEnvelope envelope)
    {
        if (envelope.Source is not null) return true;
        if (envelope.OpaqueOpc is { } opaque &&
            (opaque.SourcePackage is not null || opaque.Parts.Count > 0 || opaque.PackageRelationships.Count > 0))
            return true;

        var workbook = envelope.Workbook;
        if (workbook.Theme?.Source is not null ||
            workbook.Calculation?.Source is not null ||
            workbook.View?.Source is not null ||
            workbook.AdditionalViews.Any(view => view.Source is not null) ||
            workbook.Connections.Any(connection => connection.Source is not null) ||
            workbook.DefinedNames.Any(name => name.Source is not null))
            return true;

        return workbook.Worksheets.Any(worksheet =>
            worksheet.Source is not null ||
            worksheet.Protection?.Source is not null ||
            worksheet.Images.Any(image => image.Source is not null) ||
            worksheet.Charts.Any(chart => chart.Source is not null) ||
            worksheet.Tables.Any(table => table.Source is not null || table.QueryTable?.Source is not null) ||
            worksheet.SparklineGroups.Any(group => group.Source is not null) ||
            worksheet.PivotTables.Any(pivot => pivot.Source is not null));
    }

    internal static XlsxImportResult Import(byte[] bytes, EffectiveCodecLimits limits)
    {
        var opaque = PackageGuards.ValidateAndCollectOpaque(bytes, limits, OpcPackageProfile.Xlsx);
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var workbookPart = document.WorkbookPart ?? throw new CodecException("missing_workbook_part", "XLSX package has no Workbook part.", "xl/workbook.xml");
        var workbookRoot = workbookPart.Workbook ?? throw new CodecException("missing_workbook_root", "XLSX package has no Workbook root element.", "xl/workbook.xml");
        var dynamicArrays = new XlsxDynamicArrayCodec(workbookPart);
        var imageAssets = new XlsxImageAssetCatalog(null, limits);
        var drawings = new XlsxDrawingCodec(imageAssets);
        var workbook = new WorkbookArtifact
        {
            Id = "workbook/1",
            DateSystem = workbookRoot.WorkbookProperties?.Date1904?.Value == true ? WorkbookDateSystem._1904 : WorkbookDateSystem._1900,
        };
        var theme = new XlsxThemeCodec(workbookPart);
        if (theme.Read() is { } importedTheme) workbook.Theme = importedTheme;
        var connections = new XlsxConnectionCodec(workbookPart);
        workbook.Connections.Add(connections.Read());
        var diagnostics = new List<Diagnostic>();
        var opaqueCount = opaque.Parts.Count + opaque.PackageRelationships.Count;
        if (opaqueCount > 0)
            diagnostics.Add(CodecProtocol.Warning("opaque_content_retained", $"Retained {opaqueCount} opaque or residual OPC parts or relationships for source-bound, fail-closed export from the validated package snapshot.", opaque.Parts.FirstOrDefault()?.Path ?? opaque.PackageRelationships.FirstOrDefault()?.SourcePath));
        var sharedStrings = ReadSharedStrings(workbookPart.SharedStringTablePart);
        var styles = new XlsxCellStyleCodec(workbookPart);
        var worksheetFeatures = new XlsxWorksheetFeatureCodec(styles);
        var sheets = workbookRoot.Sheets?.Elements<Sheet>().ToArray() ?? [];
        if ((uint)sheets.Length > limits.MaxSheets)
            throw new CodecException("sheet_budget_exceeded", $"XLSX workbook has {sheets.Length} sheets and exceeds max_sheets ({limits.MaxSheets}).");
        var definedNames = new XlsxDefinedNameCodec(workbookPart, sheets.Select((sheet, index) => sheet.Name?.Value ?? $"Sheet{index + 1}").ToArray());
        var worksheetMetadata = new XlsxWorksheetMetadataCodec(workbookPart);
        workbook.DefinedNames.Add(definedNames.Read());
        var calculation = new XlsxCalculationCodec(workbookPart);
        if (calculation.Read() is { } importedCalculation) workbook.Calculation = importedCalculation;

        ulong cellCount = 0;
        var worksheetBindings = new List<(WorksheetPart Part, WorksheetArtifact Artifact)>();
        for (var index = 0; index < sheets.Length; index++)
        {
            var sheet = sheets[index];
            if (sheet.Id?.Value is not { Length: > 0 } relationshipId || workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
                throw new CodecException("missing_worksheet_part", $"Worksheet {sheet.Name?.Value ?? index.ToString(CultureInfo.InvariantCulture)} has no readable Worksheet part.");
            var target = ReadWorksheet(worksheetPart, sheet.Name?.Value ?? $"Sheet{index + 1}", index, sharedStrings, styles, dynamicArrays, diagnostics, ref cellCount, limits);
            if (new XlsxWorksheetProtectionCodec(worksheetPart).Read() is { } protection) target.Protection = protection;
            worksheetFeatures.ReadRules(worksheetPart.Worksheet!, target);
            worksheetMetadata.ReadInto(target, index);
            var tables = new XlsxTableCodec(worksheetPart, workbookPart, styles, connections);
            target.Tables.Add(tables.Read());
            target.Images.Add(drawings.Read(worksheetPart, target.Id));
            target.Charts.Add(new XlsxChartCodec().Read(worksheetPart, target.Id));
            target.SparklineGroups.Add(new XlsxSparklineCodec().Read(worksheetPart, target.Id));
            target.PivotTables.Add(new XlsxPivotTableCodec(workbookPart).Read(worksheetPart, target.Id));
            workbook.Worksheets.Add(target);
            worksheetBindings.Add((worksheetPart, target));
        }
        worksheetFeatures.ReadThreadedComments(workbookPart, worksheetBindings);
        var workbookView = new XlsxWorkbookViewCodec(workbookPart, workbook.Worksheets);
        var importedViews = workbookView.Read();
        if (importedViews.Length > 0)
        {
            workbook.View = importedViews[0];
            workbook.AdditionalViews.Add(importedViews.Skip(1));
        }

        var envelope = new ArtifactEnvelope
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Family = ArtifactFamily.Workbook,
            Workbook = workbook,
            OpaqueOpc = opaque,
            Source = new SourceIdentity
            {
                Format = "xlsx",
                PackageSha256 = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(),
                Producer = "open-office-artifact-tool/OpenChestnut",
            },
        };
        envelope.Assets.Add(imageAssets.ImportedAssets);
        envelope.Diagnostics.Add(diagnostics);
        return new XlsxImportResult(envelope, diagnostics);
    }

    private static XlsxExportResult ExportPreservingSource(ArtifactEnvelope envelope, EffectiveCodecLimits limits, int opaqueCount)
    {
        var sourceBytes = PackageGuards.ValidateSourcePackage(envelope.OpaqueOpc, envelope.Source, limits, OpcPackageProfile.Xlsx);
        if (SourceBoundWorkbookMatchesValidatedSource(envelope, sourceBytes, limits))
        {
            ValidateOutputBudget(sourceBytes, limits);
            ValidateOffice2021(sourceBytes);
            return new XlsxExportResult(sourceBytes, SourcePreservationDiagnostics(opaqueCount));
        }
        var imageAssets = new XlsxImageAssetCatalog(envelope.Assets, limits);
        var ownsTheme = false;
        var threadedRelationshipsDirty = false;
        string? themePartPath = null;
        var dirtyModeledPartPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using var stream = new MemoryStream();
        stream.Write(sourceBytes);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, isEditable: true, new OpenSettings { AutoSave = false }))
        {
            var workbookPart = document.WorkbookPart ?? throw new CodecException("missing_workbook_part", "Source XLSX package has no Workbook part.", "xl/workbook.xml");
            var workbookRoot = workbookPart.Workbook ?? throw new CodecException("missing_workbook_root", "Source XLSX package has no Workbook root element.", "xl/workbook.xml");
            var originalWorkbookXml = workbookRoot.OuterXml;
            var dynamicArrays = new XlsxDynamicArrayCodec(workbookPart);
            var sheets = workbookRoot.Sheets?.Elements<Sheet>().ToArray() ?? [];
            if (sheets.Length != envelope.Workbook.Worksheets.Count)
                throw new CodecException("source_package_topology_changed", "Source-preserving XLSX export currently requires the imported worksheet count to remain unchanged.");
            var originalWorksheetXml = new Dictionary<WorksheetPart, string>();
            foreach (var sheet in sheets)
            {
                if (sheet.Id?.Value is not { Length: > 0 } relationshipId || workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart || worksheetPart.Worksheet is null)
                    throw new CodecException("missing_worksheet_part", $"Source worksheet {sheet.Name?.Value ?? "unknown"} has no readable Worksheet part.");
                originalWorksheetXml.Add(worksheetPart, worksheetPart.Worksheet.OuterXml);
            }
            var sourceSheetNames = sheets.Select((sheet, index) => sheet.Name?.Value ?? $"Sheet{index + 1}").ToArray();
            var targetSheetNames = envelope.Workbook.Worksheets.Select(sheet => sheet.Name).ToArray();
            var definedNames = new XlsxDefinedNameCodec(workbookPart, sourceSheetNames);
            var calculation = new XlsxCalculationCodec(workbookPart);
            var worksheetMetadata = new XlsxWorksheetMetadataCodec(workbookPart);
            var workbookView = new XlsxWorkbookViewCodec(workbookPart, envelope.Workbook.Worksheets);
            worksheetMetadata.Apply(envelope.Workbook.Worksheets);
            workbookView.Apply(envelope.Workbook.View, envelope.Workbook.AdditionalViews, sourceBound: true, envelope.Workbook.Worksheets);

            var sourceDateSystem = workbookRoot.WorkbookProperties?.Date1904?.Value == true
                ? WorkbookDateSystem._1904
                : WorkbookDateSystem._1900;
            if (sourceDateSystem != envelope.Workbook.DateSystem)
            {
                if (workbookRoot.WorkbookProperties is null)
                    workbookRoot.WorkbookProperties = new WorkbookProperties();
                workbookRoot.WorkbookProperties.Date1904 = envelope.Workbook.DateSystem == WorkbookDateSystem._1904;
            }
            var sharedStrings = ReadSharedStrings(workbookPart.SharedStringTablePart);
            var theme = new XlsxThemeCodec(workbookPart);
            theme.Apply(envelope.Workbook.Theme, sourceBound: true);
            var styles = new XlsxCellStyleCodec(workbookPart);
            var worksheetFeatures = new XlsxWorksheetFeatureCodec(styles);
            var connections = new XlsxConnectionCodec(workbookPart);
            connections.Apply(envelope.Workbook.Connections, sourceBound: true);
            var drawings = new XlsxDrawingCodec(imageAssets);
            var charts = new XlsxChartCodec();
            var sparklines = new XlsxSparklineCodec();
            var nextTableId = 1U;
            var worksheetBindings = new List<(WorksheetPart Part, WorksheetArtifact Artifact)>();
            for (var index = 0; index < sheets.Length; index++)
            {
                var sheet = sheets[index];
                var source = envelope.Workbook.Worksheets[index];
                if (sheet.Id?.Value is not { Length: > 0 } relationshipId || workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
                    throw new CodecException("missing_worksheet_part", $"Source worksheet {sheet.Name?.Value ?? index.ToString(CultureInfo.InvariantCulture)} has no readable Worksheet part.");
                var originalWorksheetXmlSha256 = XlsxSparklineCodec.WorksheetXmlSha256(worksheetPart);
                var protection = new XlsxWorksheetProtectionCodec(worksheetPart);
                PatchWorksheet(worksheetPart, source, sharedStrings, styles, dynamicArrays);
                protection.Apply(source.Protection, sourceBound: true);
                worksheetFeatures.ApplyRules(worksheetPart.Worksheet!, source, sourceBound: true);
                var tables = new XlsxTableCodec(worksheetPart, workbookPart, styles, connections);
                tables.Apply(source.Tables, sourceBound: true, ref nextTableId);
                tables.Save();
                dirtyModeledPartPaths.UnionWith(tables.DirtyPartPaths);
                var originalDrawingXmlSha256 = XlsxDrawingCodec.DrawingXmlSha256(worksheetPart);
                drawings.Apply(worksheetPart, source.Id, source.Images, sourceBound: true, originalDrawingXmlSha256);
                dirtyModeledPartPaths.UnionWith(drawings.DirtyPartPaths);
                charts.Apply(worksheetPart, source.Id, source.Charts, sourceBound: true, originalDrawingXmlSha256);
                dirtyModeledPartPaths.UnionWith(charts.DirtyPartPaths);
                sparklines.Apply(worksheetPart, source.Id, source.SparklineGroups, sourceBound: true, originalWorksheetXmlSha256);
                dirtyModeledPartPaths.UnionWith(sparklines.DirtyPartPaths);
                worksheetBindings.Add((worksheetPart, source));
            }
            new XlsxPivotTableCodec(workbookPart).Apply(worksheetBindings, sourceBound: true);
            worksheetFeatures.ApplyThreadedComments(workbookPart, worksheetBindings, sourceBound: true);
            dirtyModeledPartPaths.UnionWith(worksheetFeatures.DirtyPartPaths);
            threadedRelationshipsDirty = worksheetFeatures.ThreadedRelationshipGraphDirty;
            definedNames.Apply(envelope.Workbook.DefinedNames, sourceBound: true, targetSheetNames);
            calculation.Apply(envelope.Workbook.Calculation, sourceBound: true);
            connections.Save();
            if (connections.Dirty) dirtyModeledPartPaths.Add(connections.Path);
            theme.Save();
            ownsTheme = theme.OwnsOpaqueTheme;
            themePartPath = theme.PartPath;
            styles.Save();
            foreach (var (worksheetPart, sourceXml) in originalWorksheetXml)
                if (!string.Equals(worksheetPart.Worksheet!.OuterXml, sourceXml, StringComparison.Ordinal))
                    worksheetPart.Worksheet.Save();
            if (!string.Equals(workbookRoot.OuterXml, originalWorkbookXml, StringComparison.Ordinal))
                workbookRoot.Save();
        }

        var bytes = stream.ToArray();
        ValidateOutputBudget(bytes, limits);
        ValidateOffice2021(bytes);
        var outputOpaque = PackageGuards.ValidateAndCollectOpaque(bytes, limits, OpcPackageProfile.Xlsx, includeSourcePackage: false);
        PackageGuards.AssertOpaqueGraphMatches(
            envelope.OpaqueOpc,
            outputOpaque,
            "opaque_content_not_preserved",
            ignoreRelationship: ownsTheme || threadedRelationshipsDirty
                ? item => ownsTheme && XlsxThemeCodec.IsThemeRelationship(item) || threadedRelationshipsDirty && XlsxWorksheetFeatureCodec.IsThreadedRelationship(item)
                : null,
            ignorePart: ownsTheme || dirtyModeledPartPaths.Count > 0
                ? item => (ownsTheme && themePartPath is not null && item.Path.Equals(themePartPath, StringComparison.OrdinalIgnoreCase)) || dirtyModeledPartPaths.Contains(item.Path)
                : null);
        return new XlsxExportResult(bytes, SourcePreservationDiagnostics(opaqueCount));
    }

    // The JavaScript model deliberately has no package writer. Re-import the
    // already hash-validated source and compare its complete modeled payload
    // plus owned assets before opening an editable SDK package. This removes
    // ZIP/XML serializer churn only for a genuine no-op; every semantic edit
    // remains on the existing source-bound patch path.
    private static bool SourceBoundWorkbookMatchesValidatedSource(ArtifactEnvelope requested, byte[] sourceBytes, EffectiveCodecLimits limits) =>
        SourceBoundWorkbookFingerprint(requested).SequenceEqual(
            SourceBoundWorkbookFingerprint(Import(sourceBytes, limits).Artifact));

    private static byte[] SourceBoundWorkbookFingerprint(ArtifactEnvelope envelope)
    {
        var projection = new ArtifactEnvelope
        {
            ProtocolVersion = envelope.ProtocolVersion,
            Family = envelope.Family,
            Workbook = envelope.Workbook.Clone(),
        };
        projection.Assets.Add(envelope.Assets.Select(asset => asset.Clone()));
        return SHA256.HashData(projection.ToByteArray());
    }

    private static IReadOnlyList<Diagnostic> SourcePreservationDiagnostics(int opaqueCount) =>
    [
        CodecProtocol.Warning("opaque_content_preserved", $"Preserved {opaqueCount} unsupported OPC parts or relationships from the validated source package while updating modeled workbook content."),
    ];

    private static void ValidateOutputBudget(byte[] bytes, EffectiveCodecLimits limits)
    {
        if ((ulong)bytes.LongLength > limits.MaxInputBytes)
            throw new CodecException("output_budget_exceeded", $"Generated XLSX has {bytes.LongLength} bytes and exceeds max_input_bytes ({limits.MaxInputBytes}).");
    }

    private static void PatchWorksheet(WorksheetPart worksheetPart, WorksheetArtifact source, IReadOnlyList<string> sharedStrings, XlsxCellStyleCodec styles, XlsxDynamicArrayCodec dynamicArrays)
    {
        var worksheet = worksheetPart.Worksheet ?? throw new CodecException("missing_worksheet_root", $"Worksheet {source.Name} has no Worksheet root element.");
        var formulas = XlsxFormulaCodec.ForWorksheet(worksheet, source.Name, dynamicArrays);
        PatchSheetView(worksheet, source);
        PatchColumnDimensions(worksheet, source);
        PatchRowsAndCells(worksheet, source, sharedStrings, styles, formulas, dynamicArrays);
        PatchMergedRanges(worksheet, source);
        PatchWorksheetSortState(worksheet, source, styles);
    }

    private static void ValidateOffice2021(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document).Take(2).ToArray();
        if (errors.Length == 0) return;
        var first = errors[0];
        var suffix = errors.Length > 1 ? " Additional validation errors were omitted." : string.Empty;
        throw new CodecException(
            "openxml_validation_failed",
            $"Open XML SDK Office 2021 validation failed: {first.Description}.{suffix}",
            first.Part?.Uri.ToString());
    }

    private static void PatchSheetView(Worksheet worksheet, WorksheetArtifact source)
    {
        var sheetViews = worksheet.SheetViews;
        var freeze = source.FreezePane;
        var wantsFreeze = freeze is not null && (freeze.Rows > 0 || freeze.Columns > 0);
        if (sheetViews is null)
        {
            if (source.ShowGridLines && !wantsFreeze) return;
            sheetViews = new SheetViews();
            var before = worksheet.Elements().FirstOrDefault(item => item is SheetFormatProperties or Columns or SheetData);
            if (before is null) worksheet.Append(sheetViews);
            else worksheet.InsertBefore(sheetViews, before);
        }
        var sheetView = sheetViews.Elements<SheetView>().FirstOrDefault();
        if (sheetView is null)
        {
            if (source.ShowGridLines && !wantsFreeze) return;
            sheetView = new SheetView { WorkbookViewId = 0U };
            sheetViews.Append(sheetView);
        }
        var sourceShowGridLines = sheetView.ShowGridLines?.Value ?? true;
        if (sourceShowGridLines != source.ShowGridLines) sheetView.ShowGridLines = source.ShowGridLines;
        var sourcePane = sheetView.Elements<Pane>().FirstOrDefault(item =>
            item.State?.Value == PaneStateValues.Frozen || item.State?.Value == PaneStateValues.FrozenSplit);
        if (FreezePaneMatches(sourcePane, freeze)) return;
        foreach (var pane in sheetView.Elements<Pane>().ToArray()) pane.Remove();
        if (wantsFreeze)
        {
            var desiredFreeze = freeze!;
            sheetView.InsertAt(new Pane
            {
                State = PaneStateValues.Frozen,
                HorizontalSplit = desiredFreeze.Columns,
                VerticalSplit = desiredFreeze.Rows,
                TopLeftCell = string.IsNullOrWhiteSpace(desiredFreeze.TopLeftCell) ? CellReference(desiredFreeze.Rows, desiredFreeze.Columns) : desiredFreeze.TopLeftCell,
                ActivePane = desiredFreeze.Rows > 0 && desiredFreeze.Columns > 0
                    ? PaneValues.BottomRight
                    : desiredFreeze.Rows > 0 ? PaneValues.BottomLeft : PaneValues.TopRight,
            }, 0);
        }
    }

    private static bool FreezePaneMatches(Pane? current, FreezePane? desired)
    {
        var wantsFreeze = desired is not null && (desired.Rows > 0 || desired.Columns > 0);
        if (!wantsFreeze) return current is null;
        if (current is null) return false;
        var expectedTopLeft = string.IsNullOrWhiteSpace(desired!.TopLeftCell)
            ? CellReference(desired.Rows, desired.Columns)
            : desired.TopLeftCell;
        var expectedPane = desired.Rows > 0 && desired.Columns > 0
            ? PaneValues.BottomRight
            : desired.Rows > 0 ? PaneValues.BottomLeft : PaneValues.TopRight;
        return Math.Max(0, current.VerticalSplit?.Value ?? 0) == desired.Rows &&
            Math.Max(0, current.HorizontalSplit?.Value ?? 0) == desired.Columns &&
            string.Equals(current.TopLeftCell?.Value ?? string.Empty, expectedTopLeft, StringComparison.OrdinalIgnoreCase) &&
            current.ActivePane?.Value == expectedPane;
    }

    private static void PatchColumnDimensions(Worksheet worksheet, WorksheetArtifact source)
    {
        var expected = source.ColumnDimensions.OrderBy(item => item.Column).ToArray();
        var current = worksheet.Elements<Columns>().SelectMany(item => item.Elements<Column>())
            .SelectMany(column => Enumerable.Range(
                checked((int)(column.Min?.Value ?? 1) - 1),
                checked((int)((column.Max?.Value ?? column.Min?.Value ?? 1) - (column.Min?.Value ?? 1) + 1)))
                .Select(number => new ColumnDimension
                {
                    Column = checked((uint)number),
                    Width = column.Width?.Value ?? 0,
                    Hidden = column.Hidden?.Value ?? false,
                    BestFit = column.BestFit?.Value ?? false,
                }))
            .OrderBy(item => item.Column)
            .ToArray();
        if (current.Length == expected.Length && current.Zip(expected).All(pair =>
            pair.First.Column == pair.Second.Column &&
            Math.Abs(pair.First.Width - pair.Second.Width) < 0.0000001 &&
            pair.First.Hidden == pair.Second.Hidden &&
            pair.First.BestFit == pair.Second.BestFit)) return;

        foreach (var columns in worksheet.Elements<Columns>().ToArray()) columns.Remove();
        if (expected.Length == 0) return;
        var replacement = new Columns();
        foreach (var dimension in expected)
            replacement.Append(new Column
            {
                Min = dimension.Column + 1,
                Max = dimension.Column + 1,
                Width = dimension.Width > 0 ? dimension.Width : null,
                CustomWidth = dimension.Width > 0,
                Hidden = dimension.Hidden,
                BestFit = dimension.BestFit,
            });
        var sheetData = worksheet.GetFirstChild<SheetData>();
        if (sheetData is null) worksheet.Append(replacement);
        else worksheet.InsertBefore(replacement, sheetData);
    }

    private static void PatchRowsAndCells(Worksheet worksheet, WorksheetArtifact source, IReadOnlyList<string> sharedStrings, XlsxCellStyleCodec styles, XlsxFormulaCodec formulas, XlsxDynamicArrayCodec dynamicArrays)
    {
        var sheetData = worksheet.GetFirstChild<SheetData>();
        if (sheetData is null)
        {
            sheetData = new SheetData();
            var before = worksheet.Elements().FirstOrDefault(item => item is SheetCalculationProperties or SheetProtection or ProtectedRanges or Scenarios or AutoFilter or SortState or DataConsolidate or CustomSheetViews or MergeCells or PhoneticProperties or ConditionalFormatting or DataValidations or Hyperlinks or PrintOptions or PageMargins or PageSetup or HeaderFooter or RowBreaks or ColumnBreaks or CustomProperties or CellWatches or IgnoredErrors or Drawing or LegacyDrawing or LegacyDrawingHeaderFooter or Picture or OleObjects or Controls or WebPublishItems or TableParts or WorksheetExtensionList);
            if (before is null) worksheet.Append(sheetData);
            else worksheet.InsertBefore(sheetData, before);
        }
        var rows = sheetData.Elements<Row>().ToDictionary(row => checked((uint)Math.Max(1, row.RowIndex?.Value ?? 1) - 1));
        var dimensions = source.RowDimensions.ToDictionary(item => item.Row);
        foreach (var dimension in dimensions.Values)
        {
            var row = GetOrCreateRow(sheetData, rows, dimension.Row);
            row.Height = dimension.Height > 0 ? dimension.Height : null;
            row.CustomHeight = dimension.Height > 0;
            row.Hidden = dimension.Hidden;
        }

        foreach (var sourceCell in source.Cells.OrderBy(item => item.Row).ThenBy(item => item.Column))
        {
            var row = GetOrCreateRow(sheetData, rows, sourceCell.Row);
            var reference = CellReference(sourceCell.Row, sourceCell.Column);
            var cell = row.Elements<Cell>().FirstOrDefault(item => string.Equals(item.CellReference?.Value, reference, StringComparison.OrdinalIgnoreCase));
            var current = cell is null ? null : ReadCell(cell, sourceCell.Row, sharedStrings, styles, formulas);
            if (current is not null && CellSemanticallyEqual(current, sourceCell)) continue;
            formulas.AssertCellEditable(sourceCell);
            if (cell is null)
            {
                cell = BuildCell(sourceCell, styles, dynamicArrays, sourceBound: true);
                var before = row.Elements<Cell>().FirstOrDefault(item => ParseCellReference(item.CellReference?.Value, sourceCell.Row).Column > sourceCell.Column);
                if (before is null) row.Append(cell);
                else row.InsertBefore(cell, before);
            }
            else
            {
                var replacement = BuildCell(sourceCell, styles: null);
                if (!XlsxFormulaCodec.SemanticallyEqual(current!, sourceCell)) formulas.Apply(cell, sourceCell);
                cell.CellValue = replacement.CellValue?.CloneNode(true) as CellValue;
                cell.InlineString = replacement.InlineString?.CloneNode(true) as InlineString;
                cell.DataType = replacement.DataType;
                styles.Apply(cell, sourceCell);
            }
        }
    }

    private static Row GetOrCreateRow(SheetData sheetData, IDictionary<uint, Row> rows, uint rowIndex)
    {
        if (rows.TryGetValue(rowIndex, out var row)) return row;
        row = new Row { RowIndex = rowIndex + 1 };
        var before = sheetData.Elements<Row>().FirstOrDefault(item => (item.RowIndex?.Value ?? 1) > rowIndex + 1);
        if (before is null) sheetData.Append(row);
        else sheetData.InsertBefore(row, before);
        rows[rowIndex] = row;
        return row;
    }

    private static bool CellSemanticallyEqual(CellArtifact left, CellArtifact right)
    {
        if (!XlsxFormulaCodec.SemanticallyEqual(left, right) ||
            !string.Equals(left.NumberFormatCode, right.NumberFormatCode, StringComparison.Ordinal) ||
            !Equals(left.Style, right.Style) ||
            left.ValueCase != right.ValueCase) return false;
        return left.ValueCase switch
        {
            CellArtifact.ValueOneofCase.None => true,
            CellArtifact.ValueOneofCase.StringValue => left.StringValue == right.StringValue,
            CellArtifact.ValueOneofCase.NumberValue => left.NumberValue.Equals(right.NumberValue),
            CellArtifact.ValueOneofCase.BoolValue => left.BoolValue == right.BoolValue,
            CellArtifact.ValueOneofCase.ErrorValue => left.ErrorValue == right.ErrorValue,
            _ => false,
        };
    }

    private static IReadOnlyList<string> ReadSharedStrings(SharedStringTablePart? part)
    {
        if (part is null) return [];
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        var document = XDocument.Load(reader, LoadOptions.None);
        return document.Descendants().Where(item => item.Name.LocalName == "si").Select(item => item.Value).ToArray();
    }

    private static void PatchMergedRanges(Worksheet worksheet, WorksheetArtifact source)
    {
        var expected = source.MergedRanges.OrderBy(item => item, StringComparer.OrdinalIgnoreCase).ToArray();
        var current = worksheet.Elements<MergeCells>().SelectMany(item => item.Elements<MergeCell>())
            .Select(item => item.Reference?.Value ?? string.Empty)
            .OrderBy(item => item, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (current.SequenceEqual(expected, StringComparer.OrdinalIgnoreCase)) return;
        foreach (var merges in worksheet.Elements<MergeCells>().ToArray()) merges.Remove();
        if (source.MergedRanges.Count == 0) return;
        var replacement = new MergeCells();
        foreach (var reference in source.MergedRanges) replacement.Append(new MergeCell { Reference = reference });
        var sortState = worksheet.GetFirstChild<SortState>();
        if (sortState is not null)
        {
            worksheet.InsertAfter(replacement, sortState);
            return;
        }
        var sheetData = worksheet.GetFirstChild<SheetData>();
        if (sheetData is null) worksheet.Append(replacement);
        else worksheet.InsertAfter(replacement, sheetData);
    }

    private static void PatchWorksheetSortState(Worksheet worksheet, WorksheetArtifact source, XlsxCellStyleCodec styles)
    {
        if (source.SortState is not null)
            XlsxSortStateCodec.Validate(source.SortState, null, source.Name, $"Worksheet {source.Name}", allowColumnSort: true);
        var current = worksheet.Elements<SortState>().ToArray();
        if (current.Length > 1)
        {
            if (source.SortState is null) return;
            throw new CodecException("invalid_worksheet_sort", $"Worksheet {source.Name} has multiple native sortState elements.", source.Name);
        }
        if (current.Length == 1)
        {
            var xml = XElement.Parse(current[0].OuterXml, LoadOptions.PreserveWhitespace);
            var recognized = XlsxSortStateCodec.TryRead(xml, styles, allowExtensions: true, out var currentSort);
            if (recognized)
            {
                try { XlsxSortStateCodec.Validate(currentSort!, null, source.Name, $"Worksheet {source.Name}", allowColumnSort: true); }
                catch (CodecException) { recognized = false; }
            }
            if (!recognized)
            {
                if (source.SortState is null) return;
                throw new CodecException("invalid_worksheet_sort", $"Worksheet {source.Name} has an unsupported native sortState profile.", source.Name);
            }
            if (source.SortState is null)
            {
                current[0].Remove();
                return;
            }
            XlsxSortStateCodec.Patch(xml, source.SortState, styles);
            current[0].InsertAfterSelf(new SortState(xml.ToString(SaveOptions.DisableFormatting)));
            current[0].Remove();
            return;
        }
        if (source.SortState is null) return;
        var replacement = new SortState(XlsxSortStateCodec.Create(source.SortState, styles).ToString(SaveOptions.DisableFormatting));
        var before = worksheet.Elements().FirstOrDefault(item => item is DataConsolidate or CustomSheetViews or MergeCells or PhoneticProperties or ConditionalFormatting or DataValidations or Hyperlinks or PrintOptions or PageMargins or PageSetup or HeaderFooter or RowBreaks or ColumnBreaks or CustomProperties or CellWatches or IgnoredErrors or Drawing or LegacyDrawing or LegacyDrawingHeaderFooter or Picture or OleObjects or Controls or WebPublishItems or TableParts or WorksheetExtensionList);
        if (before is null) worksheet.Append(replacement);
        else worksheet.InsertBefore(replacement, before);
    }

    private static Worksheet BuildWorksheet(WorksheetArtifact source, XlsxCellStyleCodec styles, XlsxDynamicArrayCodec dynamicArrays)
    {
        var worksheet = new Worksheet();
        var sheetView = new SheetView { WorkbookViewId = 0U, ShowGridLines = source.ShowGridLines };
        if (source.FreezePane is { } freeze && (freeze.Rows > 0 || freeze.Columns > 0))
        {
            sheetView.Append(new Pane
            {
                State = PaneStateValues.Frozen,
                HorizontalSplit = freeze.Columns,
                VerticalSplit = freeze.Rows,
                TopLeftCell = string.IsNullOrWhiteSpace(freeze.TopLeftCell) ? CellReference(freeze.Rows, freeze.Columns) : freeze.TopLeftCell,
                ActivePane = freeze.Rows > 0 && freeze.Columns > 0
                    ? PaneValues.BottomRight
                    : freeze.Rows > 0 ? PaneValues.BottomLeft : PaneValues.TopRight,
            });
        }
        worksheet.Append(new SheetViews(sheetView));

        if (source.ColumnDimensions.Count > 0)
        {
            var columns = new Columns();
            foreach (var dimension in source.ColumnDimensions.OrderBy(item => item.Column))
            {
                columns.Append(new Column
                {
                    Min = dimension.Column + 1,
                    Max = dimension.Column + 1,
                    Width = dimension.Width > 0 ? dimension.Width : null,
                    CustomWidth = dimension.Width > 0,
                    Hidden = dimension.Hidden,
                    BestFit = dimension.BestFit,
                });
            }
            worksheet.Append(columns);
        }

        var cellsByRow = source.Cells.GroupBy(cell => cell.Row).ToDictionary(group => group.Key, group => group.OrderBy(cell => cell.Column).ToArray());
        var rowDimensions = source.RowDimensions.ToDictionary(item => item.Row);
        var rowIndexes = cellsByRow.Keys.Concat(rowDimensions.Keys).Distinct().OrderBy(row => row);
        var sheetData = new SheetData();
        foreach (var rowIndex in rowIndexes)
        {
            var row = new Row { RowIndex = rowIndex + 1 };
            if (rowDimensions.TryGetValue(rowIndex, out var dimension))
            {
                if (dimension.Height > 0)
                {
                    row.Height = dimension.Height;
                    row.CustomHeight = true;
                }
                row.Hidden = dimension.Hidden;
            }
            if (cellsByRow.TryGetValue(rowIndex, out var cells))
                foreach (var cell in cells) row.Append(BuildCell(cell, styles, dynamicArrays, sourceBound: false));
            sheetData.Append(row);
        }
        worksheet.Append(sheetData);

        if (source.SortState is not null)
        {
            XlsxSortStateCodec.Validate(source.SortState, null, source.Name, $"Worksheet {source.Name}", allowColumnSort: true);
            worksheet.Append(new SortState(XlsxSortStateCodec.Create(source.SortState, styles).ToString(SaveOptions.DisableFormatting)));
        }

        if (source.MergedRanges.Count > 0)
        {
            var mergeCells = new MergeCells();
            foreach (var range in source.MergedRanges) mergeCells.Append(new MergeCell { Reference = range });
            worksheet.Append(mergeCells);
        }
        return worksheet;
    }

    private static Cell BuildCell(CellArtifact source, XlsxCellStyleCodec? styles, XlsxDynamicArrayCodec? dynamicArrays = null, bool sourceBound = false)
    {
        var cell = new Cell { CellReference = CellReference(source.Row, source.Column) };
        cell.CellFormula = XlsxFormulaCodec.Build(source);
        dynamicArrays?.ConfigureNewCell(cell, source, sourceBound);
        switch (source.ValueCase)
        {
            case CellArtifact.ValueOneofCase.StringValue:
                if (cell.CellFormula is null)
                {
                    cell.DataType = CellValues.InlineString;
                    cell.InlineString = new InlineString(new Text(source.StringValue));
                }
                else
                {
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(source.StringValue);
                }
                break;
            case CellArtifact.ValueOneofCase.NumberValue:
                if (!double.IsFinite(source.NumberValue)) throw new CodecException("non_finite_cell_value", $"Cell {cell.CellReference} has a non-finite number.");
                cell.CellValue = new CellValue(source.NumberValue.ToString("R", CultureInfo.InvariantCulture));
                break;
            case CellArtifact.ValueOneofCase.BoolValue:
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new CellValue(source.BoolValue ? "1" : "0");
                break;
            case CellArtifact.ValueOneofCase.ErrorValue:
                cell.DataType = CellValues.Error;
                cell.CellValue = new CellValue(source.ErrorValue);
                break;
        }
        styles?.Apply(cell, source);
        return cell;
    }

    private static WorksheetArtifact ReadWorksheet(
        WorksheetPart worksheetPart,
        string name,
        int index,
        IReadOnlyList<string> sharedStrings,
        XlsxCellStyleCodec styles,
        XlsxDynamicArrayCodec dynamicArrays,
        ICollection<Diagnostic> diagnostics,
        ref ulong cellCount,
        EffectiveCodecLimits limits)
    {
        var worksheet = worksheetPart.Worksheet ?? throw new CodecException("missing_worksheet_root", $"Worksheet {name} has no Worksheet root element.");
        var formulas = XlsxFormulaCodec.ForWorksheet(worksheet, name, dynamicArrays);
        foreach (var range in formulas.SourceBoundSharedFormulaRanges)
            diagnostics.Add(CodecProtocol.Warning(
                "partial_shared_formula_preserved",
                $"Worksheet {name} retains partial shared formula {range.Display}; its declared range is source-bound and read-only through the current model.",
                name,
                range.Reference));
        var target = new WorksheetArtifact { Id = $"worksheet/{index + 1}", Name = name, ShowGridLines = true };
        var view = worksheet.SheetViews?.Elements<SheetView>().FirstOrDefault();
        if (view?.ShowGridLines?.HasValue == true) target.ShowGridLines = view.ShowGridLines.Value;
        var pane = view?.Elements<Pane>().FirstOrDefault(item =>
            item.State?.Value == PaneStateValues.Frozen || item.State?.Value == PaneStateValues.FrozenSplit);
        if (pane is not null)
        {
            target.FreezePane = new FreezePane
            {
                Rows = checked((uint)Math.Max(0, pane.VerticalSplit?.Value ?? 0)),
                Columns = checked((uint)Math.Max(0, pane.HorizontalSplit?.Value ?? 0)),
                TopLeftCell = pane.TopLeftCell?.Value ?? string.Empty,
                ActivePane = pane.ActivePane?.Value.ToString() ?? string.Empty,
            };
        }

        foreach (var column in worksheet.Elements<Columns>().SelectMany(item => item.Elements<Column>()))
        {
            var min = column.Min?.Value ?? 1;
            var max = column.Max?.Value ?? min;
            if (max < min || max > 16_384 || max - min > 16_384) throw new CodecException("invalid_column_dimension", $"Worksheet {name} has invalid column span {min}:{max}.");
            for (var number = min; number <= max; number++)
            {
                target.ColumnDimensions.Add(new ColumnDimension
                {
                    Column = checked((uint)number - 1),
                    Width = column.Width?.Value ?? 0,
                    Hidden = column.Hidden?.Value ?? false,
                    BestFit = column.BestFit?.Value ?? false,
                });
            }
        }

        foreach (var row in worksheet.GetFirstChild<SheetData>()?.Elements<Row>() ?? [])
        {
            var rowIndex = checked((uint)Math.Max(1, row.RowIndex?.Value ?? 1) - 1);
            if ((row.CustomHeight?.Value ?? false) || (row.Hidden?.Value ?? false))
                target.RowDimensions.Add(new RowDimension { Row = rowIndex, Height = row.Height?.Value ?? 0, Hidden = row.Hidden?.Value ?? false });
            foreach (var cell in row.Elements<Cell>())
            {
                cellCount++;
                if (cellCount > limits.MaxCells) throw new CodecException("cell_budget_exceeded", $"XLSX workbook exceeds max_cells ({limits.MaxCells}).", name);
                target.Cells.Add(ReadCell(cell, rowIndex, sharedStrings, styles, formulas));
            }
        }
        foreach (var merge in worksheet.Elements<MergeCells>().SelectMany(item => item.Elements<MergeCell>()))
            if (merge.Reference?.Value is { Length: > 0 } reference) target.MergedRanges.Add(reference);
        var sortStates = worksheet.Elements<SortState>().ToArray();
        if (sortStates.Length == 1)
        {
            var xml = XElement.Parse(sortStates[0].OuterXml, LoadOptions.PreserveWhitespace);
            if (XlsxSortStateCodec.TryRead(xml, styles, allowExtensions: true, out var sortState))
            {
                try
                {
                    XlsxSortStateCodec.Validate(sortState!, null, name, $"Worksheet {name}", allowColumnSort: true);
                    target.SortState = sortState;
                }
                catch (CodecException)
                {
                    // Unsupported semantic geometry stays in the hash-bound
                    // source package and remains hidden from the editable wire.
                }
            }
        }
        return target;
    }

    private static CellArtifact ReadCell(Cell cell, uint fallbackRow, IReadOnlyList<string> sharedStrings, XlsxCellStyleCodec styles, XlsxFormulaCodec formulas)
    {
        var (row, column) = ParseCellReference(cell.CellReference?.Value, fallbackRow);
        var target = new CellArtifact { Row = row, Column = column, NumberFormatCode = styles.ReadNumberFormat(cell), Style = styles.ReadStyle(cell) };
        formulas.Populate(target);
        var text = cell.CellValue?.Text ?? cell.InnerText ?? string.Empty;
        var dataType = cell.DataType?.Value;
        if (dataType == CellValues.SharedString)
        {
            if (!int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out var sharedIndex) || sharedIndex < 0 || sharedIndex >= sharedStrings.Count)
                throw new CodecException("invalid_shared_string", $"Cell {cell.CellReference} references missing shared string {text}.");
            target.StringValue = sharedStrings[sharedIndex];
        }
        else if (dataType == CellValues.InlineString) target.StringValue = cell.InlineString?.InnerText ?? string.Empty;
        else if (dataType == CellValues.String) target.StringValue = text;
        else if (dataType == CellValues.Boolean) target.BoolValue = text is "1" or "true" or "TRUE";
        else if (dataType == CellValues.Error) target.ErrorValue = text;
        else if (dataType == CellValues.Date) target.StringValue = text;
        else if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)) target.NumberValue = number;
        else if (text.Length > 0) target.StringValue = text;
        return target;
    }

    private static void ValidateWorkbookBudget(WorkbookArtifact workbook, EffectiveCodecLimits limits)
    {
        XlsxThemeCodec.Validate(workbook.Theme);
        if (workbook.Worksheets.Count == 0) throw new CodecException("missing_worksheets", "Workbook artifact must contain at least one worksheet.");
        if ((uint)workbook.Worksheets.Count > limits.MaxSheets)
            throw new CodecException("sheet_budget_exceeded", $"Workbook has {workbook.Worksheets.Count} sheets and exceeds max_sheets ({limits.MaxSheets}).");
        XlsxWorksheetMetadataCodec.ValidateArtifact(workbook.Worksheets);
        XlsxWorkbookViewCodec.ValidateArtifact(workbook.View, workbook.AdditionalViews, workbook.Worksheets);
        XlsxDefinedNameCodec.ValidateArtifact(workbook.DefinedNames, workbook.Worksheets.Select(sheet => sheet.Name).ToArray());
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var tableNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        ulong cells = 0;
        foreach (var sheet in workbook.Worksheets)
        {
            if (string.IsNullOrWhiteSpace(sheet.Name) || sheet.Name.Length > 31 || sheet.Name.IndexOfAny(['[', ']', ':', '*', '?', '/', '\\']) >= 0)
                throw new CodecException("invalid_sheet_name", $"Worksheet name {sheet.Name} is invalid for XLSX.");
            if (!names.Add(sheet.Name)) throw new CodecException("duplicate_sheet_name", $"Workbook contains duplicate worksheet name {sheet.Name}.");
            if (sheet.SortState is not null)
                XlsxSortStateCodec.Validate(sheet.SortState, null, sheet.Name, $"Worksheet {sheet.Name}", allowColumnSort: true);
            XlsxWorksheetProtectionCodec.Validate(sheet.Protection);
            XlsxDrawingCodec.Validate(sheet.Images, sheet.Id);
            XlsxChartCodec.Validate(sheet.Charts, sheet.Id);
            XlsxWorksheetFeatureCodec.Validate(sheet);
            XlsxTableCodec.ValidateWorksheet(sheet);
            foreach (var table in sheet.Tables.Where(XlsxTableCodec.HasCompleteSemantics))
                if (!tableNames.Add(table.Name)) throw new CodecException("invalid_worksheet_table", $"Workbook contains duplicate table name {table.Name}.", sheet.Name);
            cells = checked(cells + (ulong)sheet.Cells.Count);
            if (cells > limits.MaxCells) throw new CodecException("cell_budget_exceeded", $"Workbook exceeds max_cells ({limits.MaxCells}).", sheet.Name);
            foreach (var cell in sheet.Cells)
            {
                if (cell.Row >= 1_048_576 || cell.Column >= 16_384) throw new CodecException("cell_out_of_range", $"Cell at row {cell.Row}, column {cell.Column} exceeds XLSX limits.", sheet.Name);
                XlsxNumberFormatCodec.Canonicalize(cell.NumberFormatCode, $"{sheet.Name}!{CellReference(cell.Row, cell.Column)}");
                XlsxCellStyleCodec.Validate(cell.Style, $"{sheet.Name}!{CellReference(cell.Row, cell.Column)}");
            }
            XlsxFormulaCodec.ValidateArtifact(sheet);
        }
    }

    private static string CellReference(uint row, uint column)
    {
        var number = checked((int)column + 1);
        Span<char> buffer = stackalloc char[3];
        var position = buffer.Length;
        while (number > 0)
        {
            number--;
            buffer[--position] = (char)('A' + number % 26);
            number /= 26;
        }
        return $"{new string(buffer[position..])}{row + 1}";
    }

    private static (uint Row, uint Column) ParseCellReference(string? reference, uint fallbackRow)
    {
        if (string.IsNullOrWhiteSpace(reference)) return (fallbackRow, 0);
        var column = 0U;
        var index = 0;
        while (index < reference.Length && char.IsLetter(reference[index]))
        {
            column = checked(column * 26 + (uint)(char.ToUpperInvariant(reference[index]) - 'A' + 1));
            index++;
        }
        if (column == 0 || !uint.TryParse(reference[index..], NumberStyles.None, CultureInfo.InvariantCulture, out var row) || row == 0)
            throw new CodecException("invalid_cell_reference", $"Cell reference {reference} is invalid.");
        return (row - 1, column - 1);
    }
}
