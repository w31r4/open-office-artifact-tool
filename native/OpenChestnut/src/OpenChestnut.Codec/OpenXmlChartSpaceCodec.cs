using System.Globalization;
using System.Xml.Linq;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns the package-agnostic DrawingML ChartSpace profile shared by XLSX and
// PPTX. Callers retain package relationships, source bindings, anchors, and
// Presentation-only combo topology; this module owns one ordinary plot,
// literal/reference caches, styling, labels, and the paired primary axes.
internal static class OpenXmlChartSpaceCodec
{
    private const int MaxSeries = 256;
    private const int MaxPoints = 1_048_576;
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    internal static bool TryRead(string xml, out SpreadsheetChartArtifact chart, out XDocument document, out bool editable)
    {
        chart = new SpreadsheetChartArtifact();
        editable = true;
        try { document = XDocument.Parse(xml, LoadOptions.PreserveWhitespace); }
        catch (System.Xml.XmlException) { document = new XDocument(); return false; }
        var root = document.Root;
        var nativeChart = root?.Element(ChartNs + "chart");
        var plotArea = nativeChart?.Element(ChartNs + "plotArea");
        if (root?.Name != ChartNs + "chartSpace" || nativeChart is null || plotArea is null || root.Element(ChartNs + "externalData") is not null) return false;
        var plots = plotArea.Elements().Where(item => item.Name == ChartNs + "barChart" || item.Name == ChartNs + "lineChart" || item.Name == ChartNs + "pieChart" || item.Name == ChartNs + "areaChart" || item.Name == ChartNs + "doughnutChart" || item.Name == ChartNs + "scatterChart" || item.Name == ChartNs + "bubbleChart").ToArray();
        if (plots.Length != 1 || plotArea.Elements().Any(item => item.Name.LocalName.EndsWith("Chart", StringComparison.Ordinal) && !plots.Contains(item))) return false;
        var plot = plots[0];
        chart.Type = plot.Name.LocalName switch
        {
            "barChart" => SpreadsheetChartType.Bar,
            "lineChart" => SpreadsheetChartType.Line,
            "pieChart" => SpreadsheetChartType.Pie,
            "areaChart" => SpreadsheetChartType.Area,
            "doughnutChart" => SpreadsheetChartType.Doughnut,
            "scatterChart" => SpreadsheetChartType.Scatter,
            "bubbleChart" => SpreadsheetChartType.Bubble,
            _ => SpreadsheetChartType.Unspecified,
        };
        if (chart.Type == SpreadsheetChartType.Unspecified) return false;
        editable &= PlotProfileEditable(plot, chart.Type);
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
            editable &= seriesEditable;
            if (!UsesNumericXAxis(chart.Type))
            {
                if (commonCategories is null) commonCategories = categories;
                else if (!commonCategories.SequenceEqual(categories, StringComparer.Ordinal)) return false;
            }
            chart.Series.Add(series);
        }
        if (!UsesNumericXAxis(chart.Type)) chart.Categories.Add(commonCategories ?? []);
        editable &= XlsxChartLineOptionsCodec.TryRead(plot, chart);
        editable &= XlsxChartDataLabelsCodec.TryRead(plot, chart);
        if (!XlsxChartAxisCodec.TryRead(plotArea, plot, chart, out var axesEditable)) editable = false;
        else editable &= axesEditable;
        return chart.Title.Length <= 32_767 && !HasControls(chart.Title) && chart.Categories.Count <= MaxPoints;
    }

    internal static XDocument Build(SpreadsheetChartArtifact chart)
    {
        var series = chart.Series.Select((item, index) => SeriesElement(item, chart.Categories, index, chart.Type)).ToArray();
        XElement plot = chart.Type switch
        {
            SpreadsheetChartType.Bar => new XElement(ChartNs + "barChart", new XElement(ChartNs + "barDir", new XAttribute("val", "col")), new XElement(ChartNs + "grouping", new XAttribute("val", "clustered")), series, XlsxChartDataLabelsCodec.Element(chart.DataLabels), new XElement(ChartNs + "axId", new XAttribute("val", "1")), new XElement(ChartNs + "axId", new XAttribute("val", "2"))),
            SpreadsheetChartType.Line => new XElement(ChartNs + "lineChart", XlsxChartLineOptionsCodec.GroupingElement(chart.LineOptions), XlsxChartLineOptionsCodec.VaryColorsElement(chart.LineOptions), series, XlsxChartDataLabelsCodec.Element(chart.DataLabels), XlsxChartLineOptionsCodec.SmoothElement(chart.LineOptions), new XElement(ChartNs + "axId", new XAttribute("val", "1")), new XElement(ChartNs + "axId", new XAttribute("val", "2"))),
            SpreadsheetChartType.Area => new XElement(ChartNs + "areaChart", new XElement(ChartNs + "grouping", new XAttribute("val", "standard")), series, XlsxChartDataLabelsCodec.Element(chart.DataLabels), new XElement(ChartNs + "axId", new XAttribute("val", "1")), new XElement(ChartNs + "axId", new XAttribute("val", "2"))),
            SpreadsheetChartType.Doughnut => new XElement(ChartNs + "doughnutChart", new XElement(ChartNs + "varyColors", new XAttribute("val", "1")), series, XlsxChartDataLabelsCodec.Element(chart.DataLabels), new XElement(ChartNs + "firstSliceAng", new XAttribute("val", "0")), new XElement(ChartNs + "holeSize", new XAttribute("val", "50"))),
            SpreadsheetChartType.Scatter => new XElement(ChartNs + "scatterChart", new XElement(ChartNs + "scatterStyle", new XAttribute("val", "marker")), new XElement(ChartNs + "varyColors", new XAttribute("val", "0")), series, XlsxChartDataLabelsCodec.Element(chart.DataLabels), new XElement(ChartNs + "axId", new XAttribute("val", "1")), new XElement(ChartNs + "axId", new XAttribute("val", "2"))),
            SpreadsheetChartType.Bubble => new XElement(ChartNs + "bubbleChart", new XElement(ChartNs + "varyColors", new XAttribute("val", "0")), series, XlsxChartDataLabelsCodec.Element(chart.DataLabels), new XElement(ChartNs + "bubble3D", new XAttribute("val", "0")), new XElement(ChartNs + "bubbleScale", new XAttribute("val", "100")), new XElement(ChartNs + "showNegBubbles", new XAttribute("val", "0")), new XElement(ChartNs + "sizeRepresents", new XAttribute("val", "area")), new XElement(ChartNs + "axId", new XAttribute("val", "1")), new XElement(ChartNs + "axId", new XAttribute("val", "2"))),
            SpreadsheetChartType.Pie => new XElement(ChartNs + "pieChart", new XElement(ChartNs + "varyColors", new XAttribute("val", "1")), series, XlsxChartDataLabelsCodec.Element(chart.DataLabels)),
            _ => throw new InvalidOperationException("Validated chart type is unsupported."),
        };
        var plotArea = new XElement(ChartNs + "plotArea", new XElement(ChartNs + "layout"), plot);
        XlsxChartAxisCodec.AppendAuthored(plotArea, chart);
        var nativeChart = new XElement(ChartNs + "chart");
        if (chart.Title.Length > 0) nativeChart.Add(XlsxChartTextStyleCodec.TitleElement(chart.Title, chart.TitleTextStyle));
        nativeChart.Add(plotArea);
        if (chart.HasLegend) nativeChart.Add(LegendElement());
        nativeChart.Add(new XElement(ChartNs + "plotVisOnly", new XAttribute("val", "1")));
        return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), new XElement(ChartNs + "chartSpace", new XAttribute(XNamespace.Xmlns + "c", ChartNs), new XAttribute(XNamespace.Xmlns + "a", DrawingNs), nativeChart));
    }

    internal static void Patch(XDocument document, SpreadsheetChartArtifact target, string errorCode, string subject)
    {
        var nativeChart = document.Root?.Element(ChartNs + "chart") ?? throw Topology(errorCode, subject, "is missing c:chart");
        PatchTitle(nativeChart, target.Title, target.TitleTextStyle, errorCode, subject);
        PatchLegend(nativeChart, target.HasLegend);
        var plotArea = nativeChart.Element(ChartNs + "plotArea") ?? throw Topology(errorCode, subject, "is missing c:plotArea");
        var plotName = target.Type switch
        {
            SpreadsheetChartType.Bar => "barChart",
            SpreadsheetChartType.Line => "lineChart",
            SpreadsheetChartType.Pie => "pieChart",
            SpreadsheetChartType.Area => "areaChart",
            SpreadsheetChartType.Doughnut => "doughnutChart",
            SpreadsheetChartType.Scatter => "scatterChart",
            SpreadsheetChartType.Bubble => "bubbleChart",
            _ => throw new InvalidOperationException("Validated chart type is unsupported."),
        };
        var plot = plotArea.Element(ChartNs + plotName) ?? throw Topology(errorCode, subject, $"is missing c:{plotName}");
        var nativeSeries = plot.Elements(ChartNs + "ser").ToArray();
        if (nativeSeries.Length != target.Series.Count) throw Topology(errorCode, subject, "series topology changed unexpectedly");
        for (var index = 0; index < nativeSeries.Length; index++) PatchSeries(nativeSeries[index], target.Series[index], target.Categories, target.Type, errorCode, subject);
        if (target.Type == SpreadsheetChartType.Line) XlsxChartLineOptionsCodec.Patch(plot, target.LineOptions);
        XlsxChartDataLabelsCodec.Patch(plot, target.DataLabels);
        XlsxChartAxisCodec.Patch(plotArea, plot, target);
    }

    internal static bool TrySeries(XElement source, SpreadsheetChartType chartType, out SpreadsheetChartSeriesArtifact series, out string[] categories, out bool editable)
    {
        series = new SpreadsheetChartSeriesArtifact(); categories = []; editable = true;
        var tx = source.Element(ChartNs + "tx");
        var directName = tx?.Element(ChartNs + "v");
        if (directName is not null) series.Name = directName.Value;
        else
        {
            var reference = tx?.Element(ChartNs + "strRef");
            var names = reference is null ? null : ReadStringPoints(reference.Element(ChartNs + "strCache"));
            if (names is null || names.Length != 1) return false;
            series.Name = names[0];
            editable = false;
        }
        if (series.Name.Length > 255 || HasControls(series.Name)) return false;
        if (UsesNumericXAxis(chartType))
        {
            var xValue = source.Element(ChartNs + "xVal");
            var yValue = source.Element(ChartNs + "yVal");
            if (xValue is null || yValue is null || !TryNumericData(xValue, out var xValues, out var xFormula) || !TryNumericData(yValue, out var values, out var valueFormula) || xValues.Length != values.Length || values.Length > MaxPoints) return false;
            series.XValueFormula = xFormula;
            series.ValueFormula = valueFormula;
            series.XValues.Add(xValues);
            series.Values.Add(values);
            if (chartType == SpreadsheetChartType.Bubble)
            {
                var bubbleSize = source.Element(ChartNs + "bubbleSize");
                if (bubbleSize is null || !TryNumericData(bubbleSize, out var bubbleSizes, out var bubbleFormula) || bubbleSizes.Length != values.Length || bubbleSizes.Any(value => value <= 0)) return false;
                series.BubbleSizeFormula = bubbleFormula;
                series.BubbleSizes.Add(bubbleSizes);
                editable &= ScalarEquals(source, "bubble3D", "0", required: false);
            }
        }
        else
        {
            var category = source.Element(ChartNs + "cat");
            var value = source.Element(ChartNs + "val");
            if (category is null || value is null || !TryStringData(category, out categories, out var categoryFormula) || !TryNumericData(value, out var values, out var valueFormula) || categories.Length != values.Length || categories.Length > MaxPoints) return false;
            series.CategoryFormula = categoryFormula;
            series.ValueFormula = valueFormula;
            series.Values.Add(values);
        }
        editable &= XlsxChartSeriesStyleCodec.TryRead(source, series);
        editable &= XlsxChartSeriesLineStyleCodec.TryRead(source, series, chartType);
        editable &= XlsxChartSeriesMarkerCodec.TryRead(source, series, chartType);
        return true;
    }

    internal static XElement SeriesElement(SpreadsheetChartSeriesArtifact series, IEnumerable<string> categories, int index, SpreadsheetChartType chartType)
    {
        var output = new XElement(ChartNs + "ser",
            new XElement(ChartNs + "idx", new XAttribute("val", index)),
            new XElement(ChartNs + "order", new XAttribute("val", index)),
            new XElement(ChartNs + "tx", new XElement(ChartNs + "v", series.Name)),
            XlsxChartSeriesStyleCodec.PropertiesElement(series, markerOnly: chartType == SpreadsheetChartType.Scatter),
            XlsxChartSeriesMarkerCodec.Element(series.Marker));
        if (UsesNumericXAxis(chartType))
        {
            output.Add(new XElement(ChartNs + "xVal", NumericData(series.XValues, series.XValueFormula)), new XElement(ChartNs + "yVal", NumericData(series.Values, series.ValueFormula)));
            if (chartType == SpreadsheetChartType.Bubble) output.Add(new XElement(ChartNs + "bubbleSize", NumericData(series.BubbleSizes, series.BubbleSizeFormula)));
        }
        else output.Add(new XElement(ChartNs + "cat", StringData(categories, series.CategoryFormula)), new XElement(ChartNs + "val", NumericData(series.Values, series.ValueFormula)));
        return output;
    }

    internal static void PatchTitle(XElement chart, string title, SpreadsheetChartTextStyleArtifact? style, string errorCode, string subject)
    {
        var existing = chart.Element(ChartNs + "title");
        if (title.Length == 0) { existing?.Remove(); return; }
        if (existing is null)
        {
            var plotArea = chart.Element(ChartNs + "plotArea") ?? throw Topology(errorCode, subject, "is missing c:plotArea");
            plotArea.AddBeforeSelf(XlsxChartTextStyleCodec.TitleElement(title, style));
            return;
        }
        var runs = existing.Descendants(DrawingNs + "t").ToArray();
        if (runs.Length == 0) throw Topology(errorCode, subject, "has a title outside the editable rich-text profile");
        runs[0].Value = title;
        foreach (var run in runs.Skip(1)) run.Value = string.Empty;
        XlsxChartTextStyleCodec.PatchTitle(existing, style);
    }

    internal static void PatchLegend(XElement chart, bool hasLegend)
    {
        var legend = chart.Element(ChartNs + "legend");
        if (!hasLegend) { legend?.Remove(); return; }
        if (legend is null) chart.Element(ChartNs + "plotArea")!.AddAfterSelf(LegendElement());
    }

    internal static void PatchSeries(XElement native, SpreadsheetChartSeriesArtifact target, IEnumerable<string> categories, SpreadsheetChartType chartType, string errorCode, string subject)
    {
        var name = native.Element(ChartNs + "tx")?.Element(ChartNs + "v") ?? throw Topology(errorCode, subject, "series name topology changed unexpectedly");
        name.Value = target.Name;
        XlsxChartSeriesStyleCodec.Patch(native, target);
        XlsxChartSeriesLineStyleCodec.Patch(native, target, markerOnly: chartType == SpreadsheetChartType.Scatter);
        XlsxChartSeriesMarkerCodec.Patch(native, target);
        if (UsesNumericXAxis(chartType))
        {
            PatchNumericData(native.Element(ChartNs + "xVal"), target.XValues, target.XValueFormula, errorCode, subject);
            PatchNumericData(native.Element(ChartNs + "yVal"), target.Values, target.ValueFormula, errorCode, subject);
            if (chartType == SpreadsheetChartType.Bubble) PatchNumericData(native.Element(ChartNs + "bubbleSize"), target.BubbleSizes, target.BubbleSizeFormula, errorCode, subject);
        }
        else
        {
            PatchStringData(native.Element(ChartNs + "cat"), categories, target.CategoryFormula, errorCode, subject);
            PatchNumericData(native.Element(ChartNs + "val"), target.Values, target.ValueFormula, errorCode, subject);
        }
    }

    internal static XElement LegendElement() => new(ChartNs + "legend", new XElement(ChartNs + "legendPos", new XAttribute("val", "r")), new XElement(ChartNs + "layout"));
    internal static bool UsesNumericXAxis(SpreadsheetChartType type) => type is SpreadsheetChartType.Scatter or SpreadsheetChartType.Bubble;

    private static bool PlotProfileEditable(XElement plot, SpreadsheetChartType type)
    {
        if (type == SpreadsheetChartType.Area) return ScalarEquals(plot, "grouping", "standard", required: true);
        if (type == SpreadsheetChartType.Doughnut) return ScalarEquals(plot, "firstSliceAng", "0", required: false) && ScalarEquals(plot, "holeSize", "50", required: true);
        if (type == SpreadsheetChartType.Scatter) return ScalarEquals(plot, "scatterStyle", "marker", required: true);
        if (type == SpreadsheetChartType.Bubble) return ScalarEquals(plot, "varyColors", "0", required: false) && ScalarEquals(plot, "bubble3D", "0", required: false) && ScalarEquals(plot, "bubbleScale", "100", required: false) && ScalarEquals(plot, "showNegBubbles", "0", required: false) && ScalarEquals(plot, "sizeRepresents", "area", required: false);
        return true;
    }

    private static bool ScalarEquals(XElement owner, string name, string expected, bool required)
    {
        var matches = owner.Elements(ChartNs + name).Take(2).ToArray();
        if (matches.Length == 0) return !required;
        if (matches.Length != 1) return false;
        var element = matches[0];
        return (string?)element.Attribute("val") == expected && !element.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != "val") && !element.Nodes().Any(node => node is XText text ? !string.IsNullOrWhiteSpace(text.Value) : true);
    }

    private static bool TryStringData(XElement source, out string[] values, out string formula)
    {
        formula = string.Empty; values = [];
        var literal = source.Element(ChartNs + "strLit");
        var reference = source.Element(ChartNs + "strRef");
        if ((literal is null) == (reference is null)) return false;
        if (reference is not null)
        {
            formula = reference.Element(ChartNs + "f")?.Value ?? string.Empty;
            if (string.IsNullOrWhiteSpace(formula) || formula.Length > 8_192 || formula.StartsWith('=') || HasControls(formula)) return false;
            values = ReadStringPoints(reference.Element(ChartNs + "strCache")) ?? [];
        }
        else values = ReadStringPoints(literal) ?? [];
        return values.All(value => value.Length <= 32_767 && !HasControls(value));
    }

    private static bool TryNumericData(XElement source, out double[] values, out string formula)
    {
        formula = string.Empty; values = [];
        var literal = source.Element(ChartNs + "numLit");
        var reference = source.Element(ChartNs + "numRef");
        if ((literal is null) == (reference is null)) return false;
        XElement? cache;
        if (reference is not null)
        {
            formula = reference.Element(ChartNs + "f")?.Value ?? string.Empty;
            if (string.IsNullOrWhiteSpace(formula) || formula.Length > 8_192 || formula.StartsWith('=') || HasControls(formula)) return false;
            cache = reference.Element(ChartNs + "numCache");
        }
        else cache = literal;
        if (cache is null || !TryOrderedPoints(cache, out var points)) return false;
        var output = new List<double>();
        foreach (var point in points)
        {
            if (!double.TryParse(point.Element(ChartNs + "v")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) || !double.IsFinite(number)) return false;
            output.Add(number);
        }
        values = output.ToArray();
        return true;
    }

    private static string[]? ReadStringPoints(XElement? source)
    {
        if (source is null || !TryOrderedPoints(source, out var points)) return null;
        return points.Select(item => item.Element(ChartNs + "v")?.Value ?? string.Empty).ToArray();
    }

    private static bool TryOrderedPoints(XElement source, out XElement[] points)
    {
        points = source.Elements(ChartNs + "pt").ToArray();
        if (points.Length > MaxPoints) return false;
        for (var index = 0; index < points.Length; index++) if ((uint?)points[index].Attribute("idx") != (uint)index) return false;
        var count = (uint?)source.Element(ChartNs + "ptCount")?.Attribute("val");
        return count is null || count.Value == points.Length;
    }

    private static XElement StringData(IEnumerable<string> values, string formula)
    {
        var cache = new XElement(ChartNs + (formula.Length > 0 ? "strCache" : "strLit"));
        AppendPoints(cache, values, value => value);
        return formula.Length > 0 ? new XElement(ChartNs + "strRef", new XElement(ChartNs + "f", formula), cache) : cache;
    }

    private static XElement NumericData(IEnumerable<double> values, string formula)
    {
        var cache = new XElement(ChartNs + (formula.Length > 0 ? "numCache" : "numLit"), new XElement(ChartNs + "formatCode", "General"));
        AppendPoints(cache, values, value => value.ToString("R", CultureInfo.InvariantCulture));
        return formula.Length > 0 ? new XElement(ChartNs + "numRef", new XElement(ChartNs + "f", formula), cache) : cache;
    }

    private static void AppendPoints<T>(XElement cache, IEnumerable<T> values, Func<T, string> format)
    {
        var array = values.ToArray();
        cache.Add(new XElement(ChartNs + "ptCount", new XAttribute("val", array.Length)));
        for (var index = 0; index < array.Length; index++) cache.Add(new XElement(ChartNs + "pt", new XAttribute("idx", index), new XElement(ChartNs + "v", format(array[index]))));
    }

    private static void PatchStringData(XElement? holder, IEnumerable<string> values, string formula, string errorCode, string subject)
    {
        if (holder is null) throw Topology(errorCode, subject, "category cache topology changed unexpectedly");
        var branch = holder.Element(ChartNs + (formula.Length > 0 ? "strRef" : "strLit")) ?? throw Topology(errorCode, subject, "category literal/reference topology changed unexpectedly");
        if (formula.Length > 0) (branch.Element(ChartNs + "f") ?? throw Topology(errorCode, subject, "category formula topology changed unexpectedly")).Value = formula;
        PatchPoints(formula.Length > 0 ? branch.Element(ChartNs + "strCache") : branch, values, value => value, errorCode, subject);
    }

    private static void PatchNumericData(XElement? holder, IEnumerable<double> values, string formula, string errorCode, string subject)
    {
        if (holder is null) throw Topology(errorCode, subject, "numeric cache topology changed unexpectedly");
        var branch = holder.Element(ChartNs + (formula.Length > 0 ? "numRef" : "numLit")) ?? throw Topology(errorCode, subject, "numeric literal/reference topology changed unexpectedly");
        if (formula.Length > 0) (branch.Element(ChartNs + "f") ?? throw Topology(errorCode, subject, "numeric formula topology changed unexpectedly")).Value = formula;
        PatchPoints(formula.Length > 0 ? branch.Element(ChartNs + "numCache") : branch, values, value => value.ToString("R", CultureInfo.InvariantCulture), errorCode, subject);
    }

    private static void PatchPoints<T>(XElement? cache, IEnumerable<T> values, Func<T, string> format, string errorCode, string subject)
    {
        if (cache is null) throw Topology(errorCode, subject, "cache topology changed unexpectedly");
        var requested = values.ToArray();
        var points = cache.Elements(ChartNs + "pt").ToArray();
        if (points.Length != requested.Length) throw Topology(errorCode, subject, "point topology changed unexpectedly");
        cache.Element(ChartNs + "ptCount")?.SetAttributeValue("val", requested.Length);
        for (var index = 0; index < points.Length; index++) (points[index].Element(ChartNs + "v") ?? throw Topology(errorCode, subject, "point value topology changed unexpectedly")).Value = format(requested[index]);
    }

    private static bool HasControls(string value) => value.Any(char.IsControl);
    private static CodecException Topology(string code, string subject, string message) => new(code, $"{subject} {message}.");
}
