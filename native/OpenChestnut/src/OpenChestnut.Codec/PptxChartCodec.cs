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
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenChestnut.Codec;

// Owns the bounded literal-data p:graphicFrame -> ChartPart projection. The
// chart semantic atoms deliberately reuse the worksheet-chart wire messages;
// PresentationML contributes only its page-relative frame.
internal static class PptxChartCodec
{
    private const string ChartGraphicDataUri = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private const int MaxSeries = 256;
    private const int MaxPoints = 1_048_576;
    private static readonly XNamespace ChartNs = ChartGraphicDataUri;
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    internal sealed record Replacement(string PartPath, string Sha256);

    internal static bool TryRead(P.GraphicFrame source, PptxPartContext context, out PresentationChart chart, out bool editable)
    {
        chart = new PresentationChart();
        editable = false;
        try
        {
            if (source.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties is not { Id.HasValue: true } ||
                source.Transform is not { } transform || !TryReadFrame(transform, out var left, out var top, out var width, out var height) ||
                source.Graphic?.GraphicData is not { } graphicData ||
                !string.Equals(graphicData.Uri?.Value, ChartGraphicDataUri, StringComparison.Ordinal) ||
                graphicData.Elements<C.ChartReference>().SingleOrDefault()?.Id?.Value is not { Length: > 0 } relationshipId)
                return false;
            ChartPart part;
            try
            {
                if (context.Owner.GetPartById(relationshipId) is not ChartPart chartPart) return false;
                part = chartPart;
            }
            catch (ArgumentOutOfRangeException)
            {
                return false;
            }
            if (!TryReadChart(ReadXml(part), out var semantic, out _, out editable)) return false;
            chart = FromSpreadsheet(semantic, left, top, width, height);
            return true;
        }
        catch (Exception error) when (error is InvalidOperationException or OverflowException or XmlException)
        {
            chart = new PresentationChart();
            editable = false;
            return false;
        }
    }

    internal static P.GraphicFrame Build(PresentationElement element, uint nativeId, SlidePart slidePart)
    {
        Validate(element.Chart, element.Id, element.Name);
        var chartPart = slidePart.AddNewPart<ChartPart>();
        WriteXml(chartPart, BuildChartDocument(ToSpreadsheet(element.Chart, element.Id, element.Name)));
        var relationshipId = slidePart.GetIdOfPart(chartPart);
        return new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = nativeId, Name = element.Name },
                new P.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.Transform(
                new A.Offset { X = element.Chart.LeftEmu, Y = element.Chart.TopEmu },
                new A.Extents { Cx = element.Chart.WidthEmu, Cy = element.Chart.HeightEmu }),
            new A.Graphic(new A.GraphicData(new C.ChartReference { Id = relationshipId }) { Uri = ChartGraphicDataUri }));
    }

    internal static Replacement Apply(P.GraphicFrame source, PresentationElement requested, PptxPartContext context)
    {
        if (!TryRead(source, context, out var original, out var editable) || !editable)
            throw new CodecException("unsupported_presentation_edit", $"Presentation chart {requested.Id} no longer matches the editable literal-data chart profile.");
        Validate(requested.Chart, requested.Id, requested.Name);
        if (requested.Chart.Type != original.Type || requested.Chart.Series.Count != original.Series.Count || requested.Chart.Categories.Count != original.Categories.Count)
            throw new CodecException("presentation_chart_topology_changed", $"Presentation chart {requested.Id} cannot change chart type, series count, or point topology.");
        for (var index = 0; index < requested.Chart.Series.Count; index++)
            if (requested.Chart.Series[index].Values.Count != original.Series[index].Values.Count)
                throw new CodecException("presentation_chart_topology_changed", $"Presentation chart {requested.Id} series {index + 1} cannot change point topology.");

        var relationshipId = source.Graphic!.GraphicData!.Elements<C.ChartReference>().Single().Id!.Value!;
        var part = (ChartPart)context.Owner.GetPartById(relationshipId);
        var document = XDocument.Parse(ReadXml(part), LoadOptions.PreserveWhitespace);
        PatchChart(document, ToSpreadsheet(requested.Chart, requested.Id, requested.Name));
        WriteXml(part, document);
        source.NonVisualGraphicFrameProperties!.NonVisualDrawingProperties!.Name = requested.Name;
        SetFrame(source.Transform!, requested.Chart);
        var bytes = ReadBytes(part);
        return new Replacement(Path(part), Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant());
    }

    internal static void Validate(PresentationChart? chart, string elementId, string name)
    {
        if (chart is null) throw Invalid(elementId, "payload is missing");
        if (string.IsNullOrWhiteSpace(name) || name.Length > 255 || HasControls(name)) throw Invalid(elementId, "name must contain 1 through 255 characters without controls");
        if (chart.LeftEmu < 0 || chart.TopEmu < 0 || chart.WidthEmu <= 0 || chart.HeightEmu <= 0)
            throw Invalid(elementId, "frame must have non-negative coordinates and positive dimensions");
        if (chart.Series.Any(series => !string.IsNullOrWhiteSpace(series.CategoryFormula) || !string.IsNullOrWhiteSpace(series.ValueFormula)))
            throw Invalid(elementId, "must use literal categories and values without workbook formulas");
        var spreadsheet = ToSpreadsheet(chart, elementId, name);
        try
        {
            XlsxChartCodec.Validate([spreadsheet], $"presentation/{elementId}");
        }
        catch (CodecException error) when (error.Code == "invalid_spreadsheet_chart")
        {
            throw Invalid(elementId, error.Message);
        }
    }

    internal static void ScrubFrame(P.GraphicFrame source)
    {
        if (source.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties is { } nonVisual) nonVisual.Name = string.Empty;
        if (source.Transform is { } transform)
        {
            transform.Offset!.X = 0L;
            transform.Offset.Y = 0L;
            transform.Extents!.Cx = 1L;
            transform.Extents.Cy = 1L;
        }
    }

    private static SpreadsheetChartArtifact ToSpreadsheet(PresentationChart source, string id, string name)
    {
        var output = new SpreadsheetChartArtifact
        {
            Id = id,
            Name = name,
            Title = source.Title,
            Type = source.Type,
            HasLegend = source.HasLegend,
            AbsoluteAnchor = new SpreadsheetAbsoluteAnchorArtifact
            {
                XEmu = source.LeftEmu,
                YEmu = source.TopEmu,
                WidthEmu = source.WidthEmu,
                HeightEmu = source.HeightEmu,
            },
        };
        output.Categories.Add(source.Categories);
        output.Series.Add(source.Series.Select(series => series.Clone()));
        if (source.XAxis is not null) output.XAxis = source.XAxis.Clone();
        if (source.YAxis is not null) output.YAxis = source.YAxis.Clone();
        if (source.DataLabels is not null) output.DataLabels = source.DataLabels.Clone();
        return output;
    }

    private static PresentationChart FromSpreadsheet(SpreadsheetChartArtifact source, long left, long top, long width, long height)
    {
        var output = new PresentationChart
        {
            LeftEmu = left,
            TopEmu = top,
            WidthEmu = width,
            HeightEmu = height,
            Type = source.Type,
            Title = source.Title,
            HasLegend = source.HasLegend,
        };
        output.Categories.Add(source.Categories);
        output.Series.Add(source.Series.Select(series => series.Clone()));
        if (source.XAxis is not null) output.XAxis = source.XAxis.Clone();
        if (source.YAxis is not null) output.YAxis = source.YAxis.Clone();
        if (source.DataLabels is not null) output.DataLabels = source.DataLabels.Clone();
        return output;
    }

    private static bool TryReadChart(string xml, out SpreadsheetChartArtifact chart, out XDocument document, out bool editable)
    {
        chart = new SpreadsheetChartArtifact();
        editable = true;
        try { document = XDocument.Parse(xml, LoadOptions.PreserveWhitespace); }
        catch (XmlException) { document = new XDocument(); return false; }
        var root = document.Root;
        var nativeChart = root?.Element(ChartNs + "chart");
        var plotArea = nativeChart?.Element(ChartNs + "plotArea");
        if (root?.Name != ChartNs + "chartSpace" || nativeChart is null || plotArea is null || root.Element(ChartNs + "externalData") is not null) return false;
        var plots = plotArea.Elements().Where(item => item.Name == ChartNs + "barChart" || item.Name == ChartNs + "lineChart" || item.Name == ChartNs + "pieChart").ToArray();
        if (plots.Length != 1 || plotArea.Elements().Any(item => item.Name.LocalName.EndsWith("Chart", StringComparison.Ordinal) && !plots.Contains(item))) return false;
        var plot = plots[0];
        chart.Type = plot.Name.LocalName switch { "barChart" => SpreadsheetChartType.Bar, "lineChart" => SpreadsheetChartType.Line, "pieChart" => SpreadsheetChartType.Pie, _ => SpreadsheetChartType.Unspecified };
        if (chart.Type == SpreadsheetChartType.Unspecified) return false;
        var title = nativeChart.Element(ChartNs + "title");
        if (title is not null)
        {
            var richText = title.Descendants(DrawingNs + "t").ToArray();
            var directValue = title.Descendants(ChartNs + "v").FirstOrDefault();
            chart.Title = richText.Length > 0 ? string.Concat(richText.Select(item => item.Value)) : directValue?.Value ?? string.Empty;
            if (richText.Length == 0) editable = false;
            editable &= XlsxChartTextStyleCodec.TryReadTitle(title, chart);
        }
        chart.HasLegend = nativeChart.Element(ChartNs + "legend") is not null;
        var nativeSeries = plot.Elements(ChartNs + "ser").ToArray();
        if (nativeSeries.Length is < 1 or > MaxSeries) return false;
        string[]? commonCategories = null;
        foreach (var native in nativeSeries)
        {
            if (!TrySeries(native, chart.Type, out var series, out var categories, out var seriesEditable)) return false;
            if (series.CategoryFormula.Length > 0 || series.ValueFormula.Length > 0) return false;
            editable &= seriesEditable;
            if (commonCategories is null) commonCategories = categories;
            else if (!commonCategories.SequenceEqual(categories, StringComparer.Ordinal)) return false;
            chart.Series.Add(series);
        }
        chart.Categories.Add(commonCategories ?? []);
        editable &= XlsxChartDataLabelsCodec.TryRead(plot, chart);
        if (!XlsxChartAxisCodec.TryRead(plotArea, plot, chart, out var axesEditable)) editable = false;
        else if (chart.Type != SpreadsheetChartType.Pie) editable &= axesEditable;
        return chart.Title.Length <= 32_767 && !HasControls(chart.Title) && chart.Categories.Count <= MaxPoints;
    }

    private static bool TrySeries(XElement source, SpreadsheetChartType type, out SpreadsheetChartSeriesArtifact series, out string[] categories, out bool editable)
    {
        series = new SpreadsheetChartSeriesArtifact(); categories = []; editable = true;
        var tx = source.Element(ChartNs + "tx");
        if (tx?.Element(ChartNs + "v") is { } directName) series.Name = directName.Value;
        else return false;
        if (series.Name.Length > 255 || HasControls(series.Name)) return false;
        var category = source.Element(ChartNs + "cat") ?? source.Element(ChartNs + "xVal");
        var value = source.Element(ChartNs + "val") ?? source.Element(ChartNs + "yVal");
        if (category is null || value is null || !TryStringData(category, out categories) || !TryNumericData(value, out var values) || categories.Length != values.Length || categories.Length > MaxPoints) return false;
        series.Values.Add(values);
        editable &= XlsxChartSeriesStyleCodec.TryRead(source, series);
        editable &= XlsxChartSeriesLineStyleCodec.TryRead(source, series);
        editable &= XlsxChartSeriesMarkerCodec.TryRead(source, series, type);
        return true;
    }

    private static bool TryStringData(XElement holder, out string[] values)
    {
        values = [];
        if (holder.Element(ChartNs + "strLit") is not { } literal || holder.Element(ChartNs + "strRef") is not null) return false;
        if (!TryOrderedPoints(literal, out var points)) return false;
        values = points.Select(item => item.Element(ChartNs + "v")?.Value ?? string.Empty).ToArray();
        return values.All(value => value.Length <= 32_767 && !HasControls(value));
    }

    private static bool TryNumericData(XElement holder, out double[] values)
    {
        values = [];
        if (holder.Element(ChartNs + "numLit") is not { } literal || holder.Element(ChartNs + "numRef") is not null || !TryOrderedPoints(literal, out var points)) return false;
        var output = new List<double>();
        foreach (var point in points)
        {
            if (!double.TryParse(point.Element(ChartNs + "v")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var value) || !double.IsFinite(value)) return false;
            output.Add(value);
        }
        values = output.ToArray();
        return true;
    }

    private static bool TryOrderedPoints(XElement source, out XElement[] points)
    {
        points = source.Elements(ChartNs + "pt").ToArray();
        if (points.Length > MaxPoints) return false;
        for (var index = 0; index < points.Length; index++) if ((uint?)points[index].Attribute("idx") != (uint)index) return false;
        var count = (uint?)source.Element(ChartNs + "ptCount")?.Attribute("val");
        return count is null || count.Value == points.Length;
    }

    private static XDocument BuildChartDocument(SpreadsheetChartArtifact chart)
    {
        var series = chart.Series.Select((item, index) => SeriesElement(item, chart.Categories, index)).ToArray();
        XElement plot = chart.Type switch
        {
            SpreadsheetChartType.Bar => new XElement(ChartNs + "barChart", new XElement(ChartNs + "barDir", new XAttribute("val", "col")), new XElement(ChartNs + "grouping", new XAttribute("val", "clustered")), series, XlsxChartDataLabelsCodec.Element(chart.DataLabels), new XElement(ChartNs + "axId", new XAttribute("val", "1")), new XElement(ChartNs + "axId", new XAttribute("val", "2"))),
            SpreadsheetChartType.Line => new XElement(ChartNs + "lineChart", new XElement(ChartNs + "grouping", new XAttribute("val", "standard")), series, XlsxChartDataLabelsCodec.Element(chart.DataLabels), new XElement(ChartNs + "axId", new XAttribute("val", "1")), new XElement(ChartNs + "axId", new XAttribute("val", "2"))),
            _ => new XElement(ChartNs + "pieChart", new XElement(ChartNs + "varyColors", new XAttribute("val", "1")), series, XlsxChartDataLabelsCodec.Element(chart.DataLabels)),
        };
        var plotArea = new XElement(ChartNs + "plotArea", new XElement(ChartNs + "layout"), plot);
        XlsxChartAxisCodec.AppendAuthored(plotArea, chart);
        var nativeChart = new XElement(ChartNs + "chart");
        if (chart.Title.Length > 0) nativeChart.Add(XlsxChartTextStyleCodec.TitleElement(chart.Title, null));
        nativeChart.Add(plotArea);
        if (chart.HasLegend) nativeChart.Add(LegendElement());
        nativeChart.Add(new XElement(ChartNs + "plotVisOnly", new XAttribute("val", "1")));
        return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), new XElement(ChartNs + "chartSpace", new XAttribute(XNamespace.Xmlns + "c", ChartNs), new XAttribute(XNamespace.Xmlns + "a", DrawingNs), nativeChart));
    }

    private static XElement SeriesElement(SpreadsheetChartSeriesArtifact series, IEnumerable<string> categories, int index) =>
        new(ChartNs + "ser",
            new XElement(ChartNs + "idx", new XAttribute("val", index)),
            new XElement(ChartNs + "order", new XAttribute("val", index)),
            new XElement(ChartNs + "tx", new XElement(ChartNs + "v", series.Name)),
            XlsxChartSeriesStyleCodec.PropertiesElement(series),
            XlsxChartSeriesMarkerCodec.Element(series.Marker),
            new XElement(ChartNs + "cat", StringData(categories)),
            new XElement(ChartNs + "val", NumericData(series.Values)));

    private static XElement StringData(IEnumerable<string> values)
    {
        var literal = new XElement(ChartNs + "strLit");
        AppendPoints(literal, values, value => value);
        return literal;
    }

    private static XElement NumericData(IEnumerable<double> values)
    {
        var literal = new XElement(ChartNs + "numLit", new XElement(ChartNs + "formatCode", "General"));
        AppendPoints(literal, values, value => value.ToString("R", CultureInfo.InvariantCulture));
        return literal;
    }

    private static void AppendPoints<T>(XElement cache, IEnumerable<T> values, Func<T, string> format)
    {
        var array = values.ToArray();
        cache.Add(new XElement(ChartNs + "ptCount", new XAttribute("val", array.Length)));
        for (var index = 0; index < array.Length; index++) cache.Add(new XElement(ChartNs + "pt", new XAttribute("idx", index), new XElement(ChartNs + "v", format(array[index]))));
    }

    private static void PatchChart(XDocument document, SpreadsheetChartArtifact target)
    {
        var nativeChart = document.Root!.Element(ChartNs + "chart")!;
        PatchTitle(nativeChart, target.Title);
        PatchLegend(nativeChart, target.HasLegend);
        var plotArea = nativeChart.Element(ChartNs + "plotArea")!;
        var plot = plotArea.Element(ChartNs + (target.Type == SpreadsheetChartType.Bar ? "barChart" : target.Type == SpreadsheetChartType.Line ? "lineChart" : "pieChart"))!;
        var nativeSeries = plot.Elements(ChartNs + "ser").ToArray();
        for (var index = 0; index < nativeSeries.Length; index++) PatchSeries(nativeSeries[index], target.Series[index], target.Categories);
        XlsxChartDataLabelsCodec.Patch(plot, target.DataLabels);
        XlsxChartAxisCodec.Patch(plotArea, plot, target);
    }

    private static void PatchTitle(XElement chart, string title)
    {
        var existing = chart.Element(ChartNs + "title");
        if (title.Length == 0) { existing?.Remove(); return; }
        if (existing is null) { chart.Element(ChartNs + "plotArea")!.AddBeforeSelf(XlsxChartTextStyleCodec.TitleElement(title, null)); return; }
        var runs = existing.Descendants(DrawingNs + "t").ToArray();
        if (runs.Length == 0) throw new CodecException("unsupported_presentation_edit", "Referenced presentation chart titles are read-only.");
        runs[0].Value = title;
        foreach (var run in runs.Skip(1)) run.Value = string.Empty;
    }

    private static void PatchLegend(XElement chart, bool hasLegend)
    {
        var legend = chart.Element(ChartNs + "legend");
        if (!hasLegend) { legend?.Remove(); return; }
        if (legend is null) chart.Element(ChartNs + "plotArea")!.AddAfterSelf(LegendElement());
    }

    private static void PatchSeries(XElement native, SpreadsheetChartSeriesArtifact target, IEnumerable<string> categories)
    {
        native.Element(ChartNs + "tx")!.Element(ChartNs + "v")!.Value = target.Name;
        XlsxChartSeriesStyleCodec.Patch(native, target);
        XlsxChartSeriesLineStyleCodec.Patch(native, target);
        XlsxChartSeriesMarkerCodec.Patch(native, target);
        PatchPoints(native.Element(ChartNs + "cat")!.Element(ChartNs + "strLit")!, categories, value => value);
        PatchPoints(native.Element(ChartNs + "val")!.Element(ChartNs + "numLit")!, target.Values, value => value.ToString("R", CultureInfo.InvariantCulture));
    }

    private static void PatchPoints<T>(XElement cache, IEnumerable<T> values, Func<T, string> format)
    {
        var requested = values.ToArray();
        var points = cache.Elements(ChartNs + "pt").ToArray();
        if (points.Length != requested.Length) throw new CodecException("presentation_chart_topology_changed", "Presentation chart point topology changed unexpectedly.");
        cache.Element(ChartNs + "ptCount")?.SetAttributeValue("val", requested.Length);
        for (var index = 0; index < points.Length; index++) points[index].Element(ChartNs + "v")!.Value = format(requested[index]);
    }

    private static XElement LegendElement() => new(ChartNs + "legend", new XElement(ChartNs + "legendPos", new XAttribute("val", "r")), new XElement(ChartNs + "layout"));

    private static bool TryReadFrame(P.Transform transform, out long left, out long top, out long width, out long height)
    {
        left = top = width = height = 0;
        if (transform.Offset?.X?.Value is null || transform.Offset.Y?.Value is null || transform.Extents?.Cx?.Value is null or <= 0 || transform.Extents.Cy?.Value is null or <= 0 || transform.Offset.X.Value < 0 || transform.Offset.Y.Value < 0) return false;
        left = transform.Offset.X.Value; top = transform.Offset.Y.Value; width = transform.Extents.Cx.Value; height = transform.Extents.Cy.Value;
        return true;
    }

    private static void SetFrame(P.Transform transform, PresentationChart chart)
    {
        transform.Offset!.X = chart.LeftEmu; transform.Offset.Y = chart.TopEmu;
        transform.Extents!.Cx = chart.WidthEmu; transform.Extents.Cy = chart.HeightEmu;
    }

    private static string ReadXml(OpenXmlPart part) => Encoding.UTF8.GetString(ReadBytes(part));
    private static byte[] ReadBytes(OpenXmlPart part) { using var stream = part.GetStream(FileMode.Open, FileAccess.Read); using var memory = new MemoryStream(); stream.CopyTo(memory); return memory.ToArray(); }
    private static void WriteXml(OpenXmlPart part, XDocument document) { using var stream = part.GetStream(FileMode.Create, FileAccess.Write); using var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false), OmitXmlDeclaration = false, Indent = false }); document.Save(writer); }
    private static string Path(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static bool HasControls(string value) => value.Any(char.IsControl);
    private static CodecException Invalid(string id, string message) => new("invalid_presentation_chart", $"Presentation chart {id} {message}.");
}
