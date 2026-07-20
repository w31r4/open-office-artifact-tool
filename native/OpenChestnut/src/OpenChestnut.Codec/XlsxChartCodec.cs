using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OpenChestnut.Codec;

// Owns the bounded worksheet GraphicFrame -> ChartPart projection. It shares a
// DrawingsPart with XlsxDrawingCodec but never interprets pictures, shapes, or
// unknown anchors. Source edits patch only hash-bound chart/name semantics and
// retain all unmodeled ChartSpace XML and relationship topology.
internal sealed class XlsxChartCodec
{
    private const long MaxEmu = 95_250_000_000L;
    private const int MaxSeries = 256;
    private const int MaxPoints = 1_048_576;
    private readonly HashSet<string> _dirtyPartPaths = new(StringComparer.OrdinalIgnoreCase);

    private sealed record ChartRecord(
        SpreadsheetChartArtifact Artifact,
        DrawingsPart DrawingPart,
        ChartPart ChartPart,
        OpenXmlCompositeElement Anchor,
        Xdr.NonVisualDrawingProperties NonVisual,
        string RelationshipId,
        int Ordinal,
        string ChartXml,
        XDocument ChartDocument);

    internal IReadOnlyCollection<string> DirtyPartPaths => _dirtyPartPaths;

    internal IReadOnlyList<SpreadsheetChartArtifact> Read(WorksheetPart worksheetPart, string worksheetId) =>
        ReadRecords(worksheetPart, worksheetId, null).Select(item => item.Artifact).ToArray();

    internal void Apply(WorksheetPart worksheetPart, string worksheetId, IReadOnlyList<SpreadsheetChartArtifact> charts, bool sourceBound, string? originalDrawingXmlSha256 = null)
    {
        Validate(charts, worksheetId);
        if (!sourceBound)
        {
            if (charts.Any(chart => chart.Source is not null))
                throw new CodecException("spreadsheet_chart_source_binding_mismatch", $"Worksheet {worksheetId} source-free charts cannot carry source bindings.");
            if (charts.Count > 0) Author(worksheetPart, charts);
            return;
        }

        var records = ReadRecords(worksheetPart, worksheetId, originalDrawingXmlSha256);
        if (records.Count != charts.Count)
            throw new CodecException("invalid_spreadsheet_chart_topology", $"Worksheet {worksheetId} source-bound chart count cannot change from {records.Count} to {charts.Count}.");
        var drawingDirty = false;
        for (var index = 0; index < records.Count; index++)
        {
            var record = records[index];
            var target = charts[index];
            ValidateBinding(target, record);
            if (!target.Id.Equals(record.Artifact.Id, StringComparison.Ordinal))
                throw new CodecException("invalid_spreadsheet_chart_topology", $"Worksheet {worksheetId} chart identity/order cannot change during source-bound export.");
            if (!AnchorSemantics(target).Equals(AnchorSemantics(record.Artifact), StringComparison.Ordinal))
                throw new CodecException("unsupported_spreadsheet_chart_edit", $"Worksheet chart {target.Id} cannot change its imported anchor geometry.", Path(record.DrawingPart));
            var changed = !SemanticHash(target).Equals(SemanticHash(record.Artifact), StringComparison.OrdinalIgnoreCase);
            if (changed && target.Source?.Editable != true)
                throw new CodecException("unsupported_spreadsheet_chart_edit", $"Worksheet chart {target.Id} is read-only because its native ChartSpace profile is outside the editable subset.", Path(record.ChartPart));
            if (!target.Name.Equals(record.Artifact.Name, StringComparison.Ordinal))
            {
                record.NonVisual.Name = target.Name;
                drawingDirty = true;
            }
            if (!changed) continue;
            PatchChart(record, target);
            _dirtyPartPaths.Add(Path(record.ChartPart));
        }
        if (!drawingDirty) return;
        var part = records.Select(item => item.DrawingPart).Distinct().Single();
        part.WorksheetDrawing!.Save();
        _dirtyPartPaths.Add(Path(part));
    }

    internal static void Validate(IEnumerable<SpreadsheetChartArtifact> charts, string worksheetId)
    {
        var ids = new HashSet<string>(StringComparer.Ordinal);
        var count = 0;
        foreach (var chart in charts)
        {
            count++;
            if (count > 1_024) throw InvalidChart(worksheetId, chart.Id, "exceeds the 1024-chart worksheet budget.");
            if (string.IsNullOrWhiteSpace(chart.Id) || chart.Id.Length > 512 || HasControls(chart.Id)) throw InvalidChart(worksheetId, chart.Id, "ID must contain 1 through 512 characters without controls.");
            if (!ids.Add(chart.Id)) throw InvalidChart(worksheetId, chart.Id, "ID must be unique within its worksheet.");
            if (string.IsNullOrWhiteSpace(chart.Name) || chart.Name.Length > 255 || HasControls(chart.Name)) throw InvalidChart(worksheetId, chart.Id, "name must contain 1 through 255 characters without controls.");
            if (chart.Title.Length > 32_767 || HasControls(chart.Title)) throw InvalidChart(worksheetId, chart.Id, "title must contain at most 32767 characters without controls.");
            if (chart.Type is not (SpreadsheetChartType.Bar or SpreadsheetChartType.Line or SpreadsheetChartType.Pie or SpreadsheetChartType.Area or SpreadsheetChartType.Doughnut or SpreadsheetChartType.Scatter or SpreadsheetChartType.Bubble)) throw InvalidChart(worksheetId, chart.Id, "type must be bar, line, pie, area, doughnut, scatter, or bubble.");
            XlsxChartAxisCodec.Validate(chart, worksheetId);
            XlsxChartTextStyleCodec.Validate(chart, worksheetId);
            XlsxChartLineOptionsCodec.Validate(chart, worksheetId);
            XlsxChartDataLabelsCodec.Validate(chart, worksheetId);
            if ((chart.Anchor is null ? 0 : 1) + (chart.TwoCellAnchor is null ? 0 : 1) + (chart.AbsoluteAnchor is null ? 0 : 1) != 1) throw InvalidChart(worksheetId, chart.Id, "must carry exactly one one-cell, two-cell, or absolute anchor.");
            ValidateAnchor(chart, worksheetId);
            if (chart.Categories.Count > MaxPoints) throw InvalidChart(worksheetId, chart.Id, $"exceeds the {MaxPoints}-category budget.");
            if (chart.Categories.Any(value => value.Length > 32_767 || HasControls(value))) throw InvalidChart(worksheetId, chart.Id, "contains a category longer than 32767 characters or with controls.");
            if (UsesNumericXAxis(chart.Type) && chart.Categories.Count != 0) throw InvalidChart(worksheetId, chart.Id, $"{ChartTypeName(chart.Type)} charts use per-series numeric x_values rather than shared text categories.");
            if (chart.Series.Count is < 1 or > MaxSeries) throw InvalidChart(worksheetId, chart.Id, $"must contain 1 through {MaxSeries} series.");
            var pointCount = 0L;
            foreach (var series in chart.Series)
            {
                if (series.Name.Length > 255 || HasControls(series.Name)) throw InvalidChart(worksheetId, chart.Id, "contains an invalid series name.");
                XlsxChartSeriesStyleCodec.Validate(series, worksheetId, chart.Id);
                if (chart.Type == SpreadsheetChartType.Scatter && series.Line is not null) throw InvalidChart(worksheetId, chart.Id, $"marker-only scatter series {series.Name} cannot carry a series line; use marker.line to style the marker border.");
                XlsxChartSeriesLineStyleCodec.Validate(series, worksheetId, chart.Id);
                XlsxChartSeriesMarkerCodec.Validate(series, chart.Type, worksheetId, chart.Id);
                if (series.CategoryFormula.Length > 8_192 || series.XValueFormula.Length > 8_192 || series.ValueFormula.Length > 8_192 || series.BubbleSizeFormula.Length > 8_192 || HasControls(series.CategoryFormula) || HasControls(series.XValueFormula) || HasControls(series.ValueFormula) || HasControls(series.BubbleSizeFormula) || series.CategoryFormula.StartsWith('=') || series.XValueFormula.StartsWith('=') || series.ValueFormula.StartsWith('=') || series.BubbleSizeFormula.StartsWith('=')) throw InvalidChart(worksheetId, chart.Id, "contains an invalid category/x-value/value/bubble-size formula.");
                if (UsesNumericXAxis(chart.Type))
                {
                    if (series.CategoryFormula.Length != 0) throw InvalidChart(worksheetId, chart.Id, $"{ChartTypeName(chart.Type)} series {series.Name} cannot carry category_formula.");
                    if (series.XValues.Count != series.Values.Count) throw InvalidChart(worksheetId, chart.Id, $"{ChartTypeName(chart.Type)} series {series.Name} has {series.XValues.Count} x values for {series.Values.Count} y values.");
                    if (chart.Type == SpreadsheetChartType.Bubble)
                    {
                        if (series.BubbleSizes.Count != series.Values.Count) throw InvalidChart(worksheetId, chart.Id, $"bubble series {series.Name} has {series.BubbleSizes.Count} sizes for {series.Values.Count} y values.");
                        if (series.BubbleSizes.Any(value => !double.IsFinite(value) || value <= 0)) throw InvalidChart(worksheetId, chart.Id, $"bubble series {series.Name} sizes must be finite and positive.");
                    }
                    else if (series.BubbleSizes.Count != 0 || series.BubbleSizeFormula.Length != 0)
                        throw InvalidChart(worksheetId, chart.Id, $"scatter series {series.Name} cannot carry bubble sizes.");
                }
                else
                {
                    if (series.XValues.Count != 0 || series.XValueFormula.Length != 0) throw InvalidChart(worksheetId, chart.Id, $"series {series.Name} x values require a scatter or bubble chart.");
                    if (series.BubbleSizes.Count != 0 || series.BubbleSizeFormula.Length != 0) throw InvalidChart(worksheetId, chart.Id, $"series {series.Name} bubble sizes require a bubble chart.");
                    if (series.Values.Count != chart.Categories.Count) throw InvalidChart(worksheetId, chart.Id, $"series {series.Name} has {series.Values.Count} values for {chart.Categories.Count} categories.");
                }
                if (series.XValues.Any(value => double.IsNaN(value) || double.IsInfinity(value)) || series.Values.Any(value => double.IsNaN(value) || double.IsInfinity(value))) throw InvalidChart(worksheetId, chart.Id, $"series {series.Name} contains a non-finite coordinate value.");
                pointCount += series.Values.Count;
                if (pointCount > MaxPoints) throw InvalidChart(worksheetId, chart.Id, $"exceeds the {MaxPoints}-value budget.");
            }
        }
    }

    private static void ValidateAnchor(SpreadsheetChartArtifact chart, string worksheetId)
    {
        if (chart.Anchor is { } oneCell)
        {
            if (oneCell.Row >= 1_048_576 || oneCell.Column >= 16_384 || oneCell.RowOffsetEmu < 0 || oneCell.RowOffsetEmu > MaxEmu || oneCell.ColumnOffsetEmu < 0 || oneCell.ColumnOffsetEmu > MaxEmu || oneCell.WidthEmu <= 0 || oneCell.WidthEmu > MaxEmu || oneCell.HeightEmu <= 0 || oneCell.HeightEmu > MaxEmu)
                throw InvalidChart(worksheetId, chart.Id, "has one-cell geometry outside bounded XLSX row/column/EMU limits.");
            return;
        }
        if (chart.TwoCellAnchor is { } twoCell)
        {
            if (twoCell.From is null || twoCell.To is null || !MarkerValid(twoCell.From) || !MarkerValid(twoCell.To) || !MarkerIsAfter(twoCell.To, twoCell.From)) throw InvalidChart(worksheetId, chart.Id, "has invalid two-cell marker geometry.");
            if (twoCell.HasEditAs && twoCell.EditAs is not (SpreadsheetTwoCellEditAs.TwoCell or SpreadsheetTwoCellEditAs.OneCell or SpreadsheetTwoCellEditAs.Absolute)) throw InvalidChart(worksheetId, chart.Id, "has an unsupported editAs value.");
            return;
        }
        var absolute = chart.AbsoluteAnchor!;
        if (absolute.XEmu < -MaxEmu || absolute.XEmu > MaxEmu || absolute.YEmu < -MaxEmu || absolute.YEmu > MaxEmu || absolute.WidthEmu <= 0 || absolute.WidthEmu > MaxEmu || absolute.HeightEmu <= 0 || absolute.HeightEmu > MaxEmu) throw InvalidChart(worksheetId, chart.Id, "has absolute geometry outside bounded signed-position/positive-extent EMU limits.");
    }

    private static bool MarkerValid(SpreadsheetCellMarkerArtifact marker) => marker.Row < 1_048_576 && marker.Column < 16_384 && marker.RowOffsetEmu is >= 0 and <= MaxEmu && marker.ColumnOffsetEmu is >= 0 and <= MaxEmu;
    private static bool MarkerIsAfter(SpreadsheetCellMarkerArtifact to, SpreadsheetCellMarkerArtifact from) =>
        (to.Column > from.Column || (to.Column == from.Column && to.ColumnOffsetEmu > from.ColumnOffsetEmu)) &&
        (to.Row > from.Row || (to.Row == from.Row && to.RowOffsetEmu > from.RowOffsetEmu));

    private void Author(WorksheetPart worksheetPart, IReadOnlyList<SpreadsheetChartArtifact> charts)
    {
        var drawingPart = EnsureDrawingPart(worksheetPart);
        var nextId = drawingPart.WorksheetDrawing!.Descendants<Xdr.NonVisualDrawingProperties>().Select(item => item.Id?.Value ?? 0U).DefaultIfEmpty(1U).Max() + 1U;
        foreach (var chart in charts)
        {
            var chartPart = drawingPart.AddNewPart<ChartPart>();
            WriteXml(chartPart, OpenXmlChartSpaceCodec.Build(chart));
            var relationshipId = drawingPart.GetIdOfPart(chartPart);
            drawingPart.WorksheetDrawing.Append(BuildAnchor(chart, relationshipId, nextId++));
        }
        drawingPart.WorksheetDrawing.Save();
    }

    private static DrawingsPart EnsureDrawingPart(WorksheetPart worksheetPart)
    {
        var worksheet = worksheetPart.Worksheet ?? throw new CodecException("missing_worksheet_root", "Worksheet chart authoring requires a Worksheet root.");
        var drawings = worksheet.Elements<S.Drawing>().ToArray();
        if (drawings.Length == 0 && worksheetPart.DrawingsPart is null)
        {
            var created = worksheetPart.AddNewPart<DrawingsPart>();
            created.WorksheetDrawing = new Xdr.WorksheetDrawing();
            var drawing = new S.Drawing { Id = worksheetPart.GetIdOfPart(created) };
            var before = worksheet.ChildElements.FirstOrDefault(item => item is S.LegacyDrawing or S.LegacyDrawingHeaderFooter or S.Picture or S.OleObjects or S.Controls or S.WebPublishItems or S.TableParts or S.ExtensionList);
            if (before is null) worksheet.Append(drawing);
            else worksheet.InsertBefore(drawing, before);
            return created;
        }
        if (drawings.Length == 1 && drawings[0].Id?.Value is { Length: > 0 } relationshipId)
        {
            try
            {
                if (worksheetPart.GetPartById(relationshipId) is DrawingsPart { WorksheetDrawing: not null } existing) return existing;
            }
            catch (ArgumentOutOfRangeException)
            {
                // Fall through to the coherent topology error below.
            }
        }
        throw new CodecException("invalid_spreadsheet_chart_topology", "Source-free worksheet chart authoring requires one coherent Drawing part.");
    }

    private IReadOnlyList<ChartRecord> ReadRecords(WorksheetPart worksheetPart, string worksheetId, string? originalDrawingXmlSha256)
    {
        var drawings = worksheetPart.Worksheet?.Elements<S.Drawing>().ToArray() ?? [];
        if (drawings.Length != 1 || drawings[0].Id?.Value is not { Length: > 0 } relationshipId) return [];
        DrawingsPart drawingPart;
        try
        {
            if (worksheetPart.GetPartById(relationshipId) is not DrawingsPart { WorksheetDrawing: not null } part) return [];
            drawingPart = part;
        }
        catch (ArgumentOutOfRangeException)
        {
            return [];
        }
        var drawingHash = originalDrawingXmlSha256 ?? Hash(drawingPart.WorksheetDrawing.OuterXml);
        var output = new List<ChartRecord>();
        for (var ordinal = 0; ordinal < drawingPart.WorksheetDrawing.ChildElements.Count; ordinal++)
        {
            if (drawingPart.WorksheetDrawing.ChildElements[ordinal] is not OpenXmlCompositeElement anchor || anchor is not Xdr.OneCellAnchor and not Xdr.TwoCellAnchor and not Xdr.AbsoluteAnchor) continue;
            if (!TryAnchor(anchor, out var oneCell, out var twoCell, out var absolute)) continue;
            if (anchor.Elements<Xdr.GraphicFrame>().SingleOrDefault() is not { } frame ||
                frame.GetFirstChild<Xdr.NonVisualGraphicFrameProperties>()?.GetFirstChild<Xdr.NonVisualDrawingProperties>() is not { Id.HasValue: true } nonVisual ||
                frame.GetFirstChild<A.Graphic>()?.GetFirstChild<A.GraphicData>() is not { } graphicData ||
                graphicData.Uri?.Value != "http://schemas.openxmlformats.org/drawingml/2006/chart" ||
                graphicData.Elements<C.ChartReference>().SingleOrDefault()?.Id?.Value is not { Length: > 0 } chartRelationshipId) continue;
            ChartPart chartPart;
            try
            {
                if (drawingPart.GetPartById(chartRelationshipId) is not ChartPart part) continue;
                chartPart = part;
            }
            catch (ArgumentOutOfRangeException)
            {
                continue;
            }
            var chartXml = ReadXml(chartPart);
            if (!OpenXmlChartSpaceCodec.TryRead(chartXml, out var chart, out var document, out var editable)) continue;
            chart.Id = $"{worksheetId}/chart/{ordinal + 1}";
            chart.Name = nonVisual.Name?.Value ?? $"Chart {ordinal + 1}";
            if (oneCell is not null) chart.Anchor = oneCell;
            else if (twoCell is not null) chart.TwoCellAnchor = twoCell;
            else chart.AbsoluteAnchor = absolute;
            if (string.IsNullOrWhiteSpace(chart.Name) || chart.Name.Length > 255 || HasControls(chart.Name)) continue;
            chart.Source = new SpreadsheetChartSourceBinding
            {
                DrawingPartPath = Path(drawingPart),
                DrawingXmlSha256 = drawingHash,
                AnchorOrdinal = checked((uint)ordinal),
                AnchorXmlSha256 = Hash(anchor.OuterXml),
                ChartPartPath = Path(chartPart),
                ChartXmlSha256 = Hash(chartXml),
                SemanticSha256 = SemanticHash(chart),
                RelationshipId = chartRelationshipId,
                NonVisualId = nonVisual.Id!.Value,
                Editable = editable,
            };
            output.Add(new ChartRecord(chart, drawingPart, chartPart, anchor, nonVisual, chartRelationshipId, ordinal, chartXml, document));
        }
        return output;
    }

    private static bool TryAnchor(OpenXmlCompositeElement anchor, out SpreadsheetOneCellAnchorArtifact? oneCell, out SpreadsheetTwoCellAnchorArtifact? twoCell, out SpreadsheetAbsoluteAnchorArtifact? absolute)
    {
        oneCell = null; twoCell = null; absolute = null;
        if (anchor is Xdr.OneCellAnchor)
        {
            var from = anchor.GetFirstChild<Xdr.FromMarker>();
            var extent = anchor.GetFirstChild<Xdr.Extent>();
            if (from is null || extent is null || !TryMarker(from, out var marker) || extent.Cx?.Value is not > 0 or > MaxEmu || extent.Cy?.Value is not > 0 or > MaxEmu) return false;
            oneCell = new SpreadsheetOneCellAnchorArtifact { Row = marker.Row, Column = marker.Column, RowOffsetEmu = marker.RowOffsetEmu, ColumnOffsetEmu = marker.ColumnOffsetEmu, WidthEmu = extent.Cx.Value, HeightEmu = extent.Cy.Value };
            return true;
        }
        if (anchor is Xdr.TwoCellAnchor nativeTwoCell)
        {
            if (!TryMarker(anchor.GetFirstChild<Xdr.FromMarker>(), out var from) || !TryMarker(anchor.GetFirstChild<Xdr.ToMarker>(), out var to) || !MarkerIsAfter(to, from)) return false;
            twoCell = new SpreadsheetTwoCellAnchorArtifact { From = from, To = to };
            if (nativeTwoCell.EditAs?.Value is { } editAs)
            {
                if (editAs == Xdr.EditAsValues.TwoCell) twoCell.EditAs = SpreadsheetTwoCellEditAs.TwoCell;
                else if (editAs == Xdr.EditAsValues.OneCell) twoCell.EditAs = SpreadsheetTwoCellEditAs.OneCell;
                else if (editAs == Xdr.EditAsValues.Absolute) twoCell.EditAs = SpreadsheetTwoCellEditAs.Absolute;
                else return false;
            }
            return true;
        }
        var position = anchor.GetFirstChild<Xdr.Position>();
        var absoluteExtent = anchor.GetFirstChild<Xdr.Extent>();
        if (position?.X?.Value is not { } x || position.Y?.Value is not { } y || x is < -MaxEmu or > MaxEmu || y is < -MaxEmu or > MaxEmu || absoluteExtent?.Cx?.Value is not > 0 or > MaxEmu || absoluteExtent.Cy?.Value is not > 0 or > MaxEmu) return false;
        absolute = new SpreadsheetAbsoluteAnchorArtifact { XEmu = x, YEmu = y, WidthEmu = absoluteExtent.Cx.Value, HeightEmu = absoluteExtent.Cy.Value };
        return true;
    }

    private static bool TryMarker(OpenXmlCompositeElement? marker, out SpreadsheetCellMarkerArtifact output)
    {
        output = new SpreadsheetCellMarkerArtifact();
        if (marker is null || !uint.TryParse(marker.GetFirstChild<Xdr.RowId>()?.Text, NumberStyles.None, CultureInfo.InvariantCulture, out var row) || row >= 1_048_576 || !uint.TryParse(marker.GetFirstChild<Xdr.ColumnId>()?.Text, NumberStyles.None, CultureInfo.InvariantCulture, out var column) || column >= 16_384 || !long.TryParse(marker.GetFirstChild<Xdr.RowOffset>()?.Text, NumberStyles.None, CultureInfo.InvariantCulture, out var rowOffset) || rowOffset < 0 || rowOffset > MaxEmu || !long.TryParse(marker.GetFirstChild<Xdr.ColumnOffset>()?.Text, NumberStyles.None, CultureInfo.InvariantCulture, out var columnOffset) || columnOffset < 0 || columnOffset > MaxEmu) return false;
        output.Row = row; output.Column = column; output.RowOffsetEmu = rowOffset; output.ColumnOffsetEmu = columnOffset;
        return true;
    }

    private static OpenXmlElement BuildAnchor(SpreadsheetChartArtifact source, string relationshipId, uint nonVisualId)
    {
        var extent = source.Anchor is { } one ? (one.WidthEmu, one.HeightEmu) : source.AbsoluteAnchor is { } absolute ? (absolute.WidthEmu, absolute.HeightEmu) : (1L, 1L);
        var frame = new Xdr.GraphicFrame(
            new Xdr.NonVisualGraphicFrameProperties(
                new Xdr.NonVisualDrawingProperties { Id = nonVisualId, Name = source.Name },
                new Xdr.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoGrouping = true })),
            new Xdr.Transform(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = extent.Item1, Cy = extent.Item2 }),
            new A.Graphic(new A.GraphicData(new C.ChartReference { Id = relationshipId }) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }));
        if (source.Anchor is { } oneCell)
            return new Xdr.OneCellAnchor(BuildFrom(oneCell.Row, oneCell.Column, oneCell.RowOffsetEmu, oneCell.ColumnOffsetEmu), new Xdr.Extent { Cx = oneCell.WidthEmu, Cy = oneCell.HeightEmu }, frame, new Xdr.ClientData());
        if (source.TwoCellAnchor is { } twoCell)
        {
            var output = new Xdr.TwoCellAnchor(BuildFrom(twoCell.From), BuildTo(twoCell.To), frame, new Xdr.ClientData());
            if (twoCell.HasEditAs) output.EditAs = twoCell.EditAs switch { SpreadsheetTwoCellEditAs.TwoCell => Xdr.EditAsValues.TwoCell, SpreadsheetTwoCellEditAs.OneCell => Xdr.EditAsValues.OneCell, SpreadsheetTwoCellEditAs.Absolute => Xdr.EditAsValues.Absolute, _ => throw new InvalidOperationException() };
            return output;
        }
        var absoluteAnchor = source.AbsoluteAnchor!;
        return new Xdr.AbsoluteAnchor(new Xdr.Position { X = absoluteAnchor.XEmu, Y = absoluteAnchor.YEmu }, new Xdr.Extent { Cx = absoluteAnchor.WidthEmu, Cy = absoluteAnchor.HeightEmu }, frame, new Xdr.ClientData());
    }

    private static Xdr.FromMarker BuildFrom(SpreadsheetCellMarkerArtifact marker) => BuildFrom(marker.Row, marker.Column, marker.RowOffsetEmu, marker.ColumnOffsetEmu);
    private static Xdr.FromMarker BuildFrom(uint row, uint column, long rowOffset, long columnOffset) => new(new Xdr.ColumnId(column.ToString(CultureInfo.InvariantCulture)), new Xdr.ColumnOffset(columnOffset.ToString(CultureInfo.InvariantCulture)), new Xdr.RowId(row.ToString(CultureInfo.InvariantCulture)), new Xdr.RowOffset(rowOffset.ToString(CultureInfo.InvariantCulture)));
    private static Xdr.ToMarker BuildTo(SpreadsheetCellMarkerArtifact marker) => new(new Xdr.ColumnId(marker.Column.ToString(CultureInfo.InvariantCulture)), new Xdr.ColumnOffset(marker.ColumnOffsetEmu.ToString(CultureInfo.InvariantCulture)), new Xdr.RowId(marker.Row.ToString(CultureInfo.InvariantCulture)), new Xdr.RowOffset(marker.RowOffsetEmu.ToString(CultureInfo.InvariantCulture)));

    private static void ValidateBinding(SpreadsheetChartArtifact target, ChartRecord record)
    {
        var source = target.Source;
        var expected = record.Artifact.Source;
        if (source is null || !source.DrawingPartPath.Equals(expected.DrawingPartPath, StringComparison.OrdinalIgnoreCase) || !source.DrawingXmlSha256.Equals(expected.DrawingXmlSha256, StringComparison.OrdinalIgnoreCase) || source.AnchorOrdinal != expected.AnchorOrdinal || !source.AnchorXmlSha256.Equals(expected.AnchorXmlSha256, StringComparison.OrdinalIgnoreCase) || !source.ChartPartPath.Equals(expected.ChartPartPath, StringComparison.OrdinalIgnoreCase) || !source.ChartXmlSha256.Equals(expected.ChartXmlSha256, StringComparison.OrdinalIgnoreCase) || !source.SemanticSha256.Equals(expected.SemanticSha256, StringComparison.OrdinalIgnoreCase) || !source.RelationshipId.Equals(expected.RelationshipId, StringComparison.Ordinal) || source.NonVisualId != expected.NonVisualId || source.Editable != expected.Editable)
            throw new CodecException("spreadsheet_chart_source_binding_mismatch", $"Worksheet chart {target.Id} does not match its hash-bound Drawing/Chart part source locator.", expected.ChartPartPath);
    }

    private static void PatchChart(ChartRecord record, SpreadsheetChartArtifact target)
    {
        if (target.Type != record.Artifact.Type || target.Series.Count != record.Artifact.Series.Count) throw new CodecException("unsupported_spreadsheet_chart_edit", $"Worksheet chart {target.Id} cannot change chart type or series topology.", Path(record.ChartPart));
        if ((target.XAxis is null) != (record.Artifact.XAxis is null) || (target.YAxis is null) != (record.Artifact.YAxis is null)) throw new CodecException("unsupported_spreadsheet_chart_edit", $"Worksheet chart {target.Id} cannot change primary-axis message topology.", Path(record.ChartPart));
        if (target.Categories.Count != record.Artifact.Categories.Count) throw new CodecException("unsupported_spreadsheet_chart_edit", $"Worksheet chart {target.Id} cannot change category/point topology.", Path(record.ChartPart));
        for (var index = 0; index < target.Series.Count; index++)
        {
            var requested = target.Series[index];
            var original = record.Artifact.Series[index];
            if (requested.Values.Count != original.Values.Count || requested.XValues.Count != original.XValues.Count || requested.BubbleSizes.Count != original.BubbleSizes.Count || string.IsNullOrEmpty(requested.CategoryFormula) != string.IsNullOrEmpty(original.CategoryFormula) || string.IsNullOrEmpty(requested.XValueFormula) != string.IsNullOrEmpty(original.XValueFormula) || string.IsNullOrEmpty(requested.ValueFormula) != string.IsNullOrEmpty(original.ValueFormula) || string.IsNullOrEmpty(requested.BubbleSizeFormula) != string.IsNullOrEmpty(original.BubbleSizeFormula)) throw new CodecException("unsupported_spreadsheet_chart_edit", $"Worksheet chart {target.Id} series {index + 1} cannot change cache count or literal/reference topology.", Path(record.ChartPart));
        }
        OpenXmlChartSpaceCodec.Patch(record.ChartDocument, target, "unsupported_spreadsheet_chart_edit", $"Worksheet chart {target.Id}");
        WriteXml(record.ChartPart, record.ChartDocument);
    }

    private static string SemanticHash(SpreadsheetChartArtifact chart) => Hash(string.Join('\0', chart.Id, chart.Name, chart.Title, XlsxChartTextStyleCodec.Semantics(chart.TitleTextStyle), XlsxChartLineOptionsCodec.Semantics(chart.LineOptions), XlsxChartDataLabelsCodec.Semantics(chart.DataLabels), ((int)chart.Type).ToString(CultureInfo.InvariantCulture), chart.HasLegend ? "1" : "0", AnchorSemantics(chart), XlsxChartAxisCodec.Semantics(chart), string.Join('\u001e', chart.Categories), string.Join('\u001d', chart.Series.Select(series => string.Join('\u001f', series.Name, series.CategoryFormula, series.XValueFormula, series.ValueFormula, series.BubbleSizeFormula, XlsxChartSeriesStyleCodec.Semantics(series), XlsxChartSeriesLineStyleCodec.Semantics(series.Line), XlsxChartSeriesMarkerCodec.Semantics(series.Marker), string.Join(',', series.XValues.Select(value => value.ToString("R", CultureInfo.InvariantCulture))), string.Join(',', series.Values.Select(value => value.ToString("R", CultureInfo.InvariantCulture))), string.Join(',', series.BubbleSizes.Select(value => value.ToString("R", CultureInfo.InvariantCulture))))))));

    private static bool UsesNumericXAxis(SpreadsheetChartType type) =>
        type is SpreadsheetChartType.Scatter or SpreadsheetChartType.Bubble;

    private static string ChartTypeName(SpreadsheetChartType type) =>
        type == SpreadsheetChartType.Bubble ? "bubble" : "scatter";

    private static string AnchorSemantics(SpreadsheetChartArtifact chart)
    {
        if (chart.Anchor is { } one) return string.Join('\0', "oneCell", one.Row, one.Column, one.RowOffsetEmu, one.ColumnOffsetEmu, one.WidthEmu, one.HeightEmu);
        if (chart.TwoCellAnchor is { } two) return string.Join('\0', "twoCell", MarkerSemantics(two.From), MarkerSemantics(two.To), two.HasEditAs ? ((int)two.EditAs).ToString(CultureInfo.InvariantCulture) : "absent");
        var absolute = chart.AbsoluteAnchor!;
        return string.Join('\0', "absolute", absolute.XEmu, absolute.YEmu, absolute.WidthEmu, absolute.HeightEmu);
    }

    private static string MarkerSemantics(SpreadsheetCellMarkerArtifact marker) => string.Join('\0', marker.Row, marker.Column, marker.RowOffsetEmu, marker.ColumnOffsetEmu);
    private static string ReadXml(OpenXmlPart part) { using var reader = new StreamReader(part.GetStream(FileMode.Open, FileAccess.Read), Encoding.UTF8, true, 4096, leaveOpen: false); return reader.ReadToEnd(); }
    private static void WriteXml(OpenXmlPart part, XDocument document) { using var stream = part.GetStream(FileMode.Create, FileAccess.Write); using var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false, OmitXmlDeclaration = false }); document.Save(writer); }
    private static string Hash(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();
    private static string Path(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static bool HasControls(string value) => value.Any(char.IsControl);
    private static CodecException InvalidChart(string worksheetId, string chartId, string message) => new("invalid_spreadsheet_chart", $"Worksheet {worksheetId} chart {chartId} {message}");
}
