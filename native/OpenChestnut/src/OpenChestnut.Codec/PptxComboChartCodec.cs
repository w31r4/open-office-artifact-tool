using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns the one deliberate mixed-plot profile: literal primary-axis clustered
// bars plus all-primary or all-secondary standard lines. Keeping this topology
// separate prevents the ordinary chart codec from accumulating type-specific
// branches while making every unsupported combo graph source-bound rather than
// reconstructed.
internal static partial class PptxChartCodec
{
    private sealed record ComboNativeSeries(SpreadsheetChartType Type, PresentationChartAxisGroup AxisGroup, XElement Element, uint Order);

    private static XDocument BuildPresentationChartDocument(PresentationChart chart, string id, string name) =>
        chart.Type == SpreadsheetChartType.Combo
            ? BuildComboChartDocument(chart, id, name)
            : BuildChartDocument(ToSpreadsheet(chart, id, name));

    private static void PatchPresentationChart(XDocument document, PresentationChart chart, string id, string name)
    {
        if (chart.Type == SpreadsheetChartType.Combo) PatchComboChart(document, chart, id, name);
        else PatchChart(document, ToSpreadsheet(chart, id, name));
    }

    private static bool PresentationChartTopologyMatches(PresentationChart requested, PresentationChart original)
    {
        if (requested.Type != original.Type || requested.Categories.Count != original.Categories.Count) return false;
        if (requested.Type != SpreadsheetChartType.Combo)
        {
            if (requested.ComboSeries.Count != 0 || original.ComboSeries.Count != 0 || requested.Series.Count != original.Series.Count) return false;
            return requested.Series.Zip(original.Series).All(pair => pair.First.Values.Count == pair.Second.Values.Count);
        }
        if (requested.Series.Count != 0 || original.Series.Count != 0 || requested.ComboSeries.Count != original.ComboSeries.Count) return false;
        if ((requested.SecondaryXAxis is null) != (original.SecondaryXAxis is null) || (requested.SecondaryYAxis is null) != (original.SecondaryYAxis is null)) return false;
        return requested.ComboSeries.Zip(original.ComboSeries).All(pair =>
            pair.First.Type == pair.Second.Type &&
            ComboAxisGroup(pair.First) == ComboAxisGroup(pair.Second) &&
            pair.First.Series is not null && pair.Second.Series is not null &&
            pair.First.Series.Values.Count == pair.Second.Series.Values.Count);
    }

    private static void ValidateComboChart(PresentationChart chart, string elementId, string name)
    {
        if (chart.Series.Count != 0) throw Invalid(elementId, "must keep series empty when type is combo");
        if (chart.ComboSeries.Count is < 2 or > MaxSeries) throw Invalid(elementId, $"must contain 2 through {MaxSeries} combo_series entries");
        if (chart.Categories.Count > MaxPoints || chart.Categories.Any(value => value.Length > 32_767 || HasControls(value))) throw Invalid(elementId, "contains invalid categories");

        var bar = new List<SpreadsheetChartSeriesArtifact>();
        var line = new List<SpreadsheetChartSeriesArtifact>();
        var hasSecondaryLine = false;
        foreach (var entry in chart.ComboSeries)
        {
            if (entry.Series is null) throw Invalid(elementId, "contains a combo series without payload");
            var axisGroup = ComboAxisGroup(entry);
            if (entry.Type == SpreadsheetChartType.Bar)
            {
                if (axisGroup != PresentationChartAxisGroup.Primary) throw Invalid(elementId, "supports bar series only on the primary axis pair");
                bar.Add(entry.Series);
            }
            else if (entry.Type == SpreadsheetChartType.Line)
            {
                hasSecondaryLine |= axisGroup == PresentationChartAxisGroup.Secondary;
                line.Add(entry.Series);
            }
            else throw Invalid(elementId, "combo series type must be bar or line");
            if (!string.IsNullOrWhiteSpace(entry.Series.CategoryFormula) || !string.IsNullOrWhiteSpace(entry.Series.ValueFormula) ||
                !string.IsNullOrWhiteSpace(entry.Series.XValueFormula) || !string.IsNullOrWhiteSpace(entry.Series.BubbleSizeFormula))
                throw Invalid(elementId, "must use literal categories and values without workbook formulas");
        }
        if (bar.Count == 0 || line.Count == 0) throw Invalid(elementId, "must contain at least one bar series and one line series");
        if (hasSecondaryLine && chart.ComboSeries.Where(entry => entry.Type == SpreadsheetChartType.Line).Any(entry => ComboAxisGroup(entry) != PresentationChartAxisGroup.Secondary))
            throw Invalid(elementId, "cannot mix primary and secondary line series");
        if (!hasSecondaryLine && (chart.SecondaryXAxis is not null || chart.SecondaryYAxis is not null))
            throw Invalid(elementId, "cannot carry secondary axes without a secondary line plot");
        if ((chart.SecondaryXAxis is null) != (chart.SecondaryYAxis is null))
            throw Invalid(elementId, "must carry both secondary axes or neither");

        foreach (var (type, series) in new[] { (SpreadsheetChartType.Bar, (IReadOnlyList<SpreadsheetChartSeriesArtifact>)bar), (SpreadsheetChartType.Line, (IReadOnlyList<SpreadsheetChartSeriesArtifact>)line) })
        {
            try { XlsxChartCodec.Validate([ComboSpreadsheetChart(chart, elementId, name, type, series)], $"presentation/{elementId}"); }
            catch (CodecException error) when (error.Code == "invalid_spreadsheet_chart") { throw Invalid(elementId, error.Message); }
        }
        if (hasSecondaryLine && chart.SecondaryXAxis is not null)
        {
            var secondaryProbe = ComboSpreadsheetChart(chart, elementId, name, SpreadsheetChartType.Line, line);
            secondaryProbe.XAxis = chart.SecondaryXAxis.Clone();
            secondaryProbe.YAxis = chart.SecondaryYAxis!.Clone();
            try { XlsxChartCodec.Validate([secondaryProbe], $"presentation/{elementId}"); }
            catch (CodecException error) when (error.Code == "invalid_spreadsheet_chart") { throw Invalid(elementId, error.Message); }
        }
    }

    private static SpreadsheetChartArtifact ComboSpreadsheetChart(PresentationChart source, string id, string name, SpreadsheetChartType type, IEnumerable<SpreadsheetChartSeriesArtifact> series)
    {
        var output = new SpreadsheetChartArtifact
        {
            Id = id,
            Name = name,
            Title = source.Title,
            Type = type,
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
        output.Series.Add(series.Select(item => item.Clone()));
        if (source.XAxis is not null) output.XAxis = source.XAxis.Clone();
        if (source.YAxis is not null) output.YAxis = source.YAxis.Clone();
        if (source.DataLabels is not null) output.DataLabels = source.DataLabels.Clone();
        return output;
    }

    private static SpreadsheetChartArtifact ComboAxisCarrier(PresentationChart source, string id, string name, bool secondary = false)
    {
        var output = ComboSpreadsheetChart(source, id, name, SpreadsheetChartType.Bar, []);
        if (secondary)
        {
            output.XAxis = source.SecondaryXAxis?.Clone();
            output.YAxis = source.SecondaryYAxis?.Clone();
        }
        return output;
    }

    private static PresentationChartAxisGroup ComboAxisGroup(PresentationComboSeriesArtifact entry) =>
        entry.AxisGroup == PresentationChartAxisGroup.Secondary
            ? PresentationChartAxisGroup.Secondary
            : PresentationChartAxisGroup.Primary;

    private static bool HasSecondaryComboLine(PresentationChart chart) =>
        chart.ComboSeries.Any(entry => entry.Type == SpreadsheetChartType.Line && ComboAxisGroup(entry) == PresentationChartAxisGroup.Secondary);

    private static bool TryReadComboChart(string xml, out PresentationChart chart, out XDocument document, out bool editable)
    {
        chart = new PresentationChart();
        editable = true;
        try { document = XDocument.Parse(xml, LoadOptions.PreserveWhitespace); }
        catch (XmlException) { document = new XDocument(); return false; }
        var root = document.Root;
        var nativeChart = root?.Element(ChartNs + "chart");
        var plotArea = nativeChart?.Element(ChartNs + "plotArea");
        if (root?.Name != ChartNs + "chartSpace" || nativeChart is null || plotArea is null || root.Element(ChartNs + "externalData") is not null) return false;
        var plots = plotArea.Elements().Where(item => item.Name.LocalName.EndsWith("Chart", StringComparison.Ordinal)).ToArray();
        if (plots.Length != 2) return false;
        var barPlot = plots.SingleOrDefault(item => item.Name == ChartNs + "barChart");
        var linePlot = plots.SingleOrDefault(item => item.Name == ChartNs + "lineChart");
        if (barPlot is null || linePlot is null || !IsCanonicalComboPlot(barPlot, SpreadsheetChartType.Bar) || !IsCanonicalComboPlot(linePlot, SpreadsheetChartType.Line)) return false;
        var lineUsesSecondaryAxes = !SharesPrimaryAxes(barPlot, linePlot);

        var title = nativeChart.Element(ChartNs + "title");
        if (title is not null)
        {
            var richText = title.Descendants(DrawingNs + "t").ToArray();
            var directValue = title.Descendants(ChartNs + "v").FirstOrDefault();
            chart.Title = richText.Length > 0 ? string.Concat(richText.Select(item => item.Value)) : directValue?.Value ?? string.Empty;
            if (richText.Length == 0) editable = false;
            var titleProbe = new SpreadsheetChartArtifact();
            editable &= XlsxChartTextStyleCodec.TryReadTitle(title, titleProbe);
        }
        chart.Type = SpreadsheetChartType.Combo;
        chart.HasLegend = nativeChart.Element(ChartNs + "legend") is not null;

        if (!TryReadComboSeries(barPlot, SpreadsheetChartType.Bar, PresentationChartAxisGroup.Primary, out var barSeries) ||
            !TryReadComboSeries(linePlot, SpreadsheetChartType.Line, lineUsesSecondaryAxes ? PresentationChartAxisGroup.Secondary : PresentationChartAxisGroup.Primary, out var lineSeries)) return false;
        var orderedSeries = barSeries.Concat(lineSeries).OrderBy(item => item.Order).ToArray();
        if (orderedSeries.Length is < 2 or > MaxSeries || orderedSeries.Select(item => item.Order).Distinct().Count() != orderedSeries.Length ||
            !orderedSeries.Select(item => item.Order).SequenceEqual(Enumerable.Range(0, orderedSeries.Length).Select(index => (uint)index))) return false;

        string[]? commonCategories = null;
        foreach (var native in orderedSeries)
        {
            if (!TrySeries(native.Element, native.Type, out var series, out var categories, out var seriesEditable) ||
                series.CategoryFormula.Length > 0 || series.ValueFormula.Length > 0 || series.XValueFormula.Length > 0 || series.BubbleSizeFormula.Length > 0) return false;
            editable &= seriesEditable;
            if (commonCategories is null) commonCategories = categories;
            else if (!commonCategories.SequenceEqual(categories, StringComparer.Ordinal)) return false;
            chart.ComboSeries.Add(new PresentationComboSeriesArtifact { Type = native.Type, AxisGroup = native.AxisGroup, Series = series });
        }
        chart.Categories.Add(commonCategories ?? []);
        if (chart.Categories.Count > MaxPoints) return false;

        var barLabels = new SpreadsheetChartArtifact();
        var lineLabels = new SpreadsheetChartArtifact();
        if (!XlsxChartDataLabelsCodec.TryRead(barPlot, barLabels) || !XlsxChartDataLabelsCodec.TryRead(linePlot, lineLabels) ||
            XlsxChartDataLabelsCodec.Semantics(barLabels.DataLabels) != XlsxChartDataLabelsCodec.Semantics(lineLabels.DataLabels)) return false;
        if (barLabels.DataLabels is not null) chart.DataLabels = barLabels.DataLabels;

        var axisCarrier = ComboAxisCarrier(chart, "combo", "combo");
        if (!XlsxChartAxisCodec.TryRead(plotArea, barPlot, axisCarrier, out var axesEditable)) return false;
        chart.XAxis = axisCarrier.XAxis;
        chart.YAxis = axisCarrier.YAxis;
        editable &= axesEditable;
        if (lineUsesSecondaryAxes)
        {
            var secondaryAxisCarrier = ComboAxisCarrier(chart, "combo", "combo", secondary: true);
            if (!XlsxChartAxisCodec.TryReadPresentationSecondary(plotArea, linePlot, secondaryAxisCarrier, out var secondaryAxesEditable)) return false;
            chart.SecondaryXAxis = secondaryAxisCarrier.XAxis;
            chart.SecondaryYAxis = secondaryAxisCarrier.YAxis;
            editable &= secondaryAxesEditable;
        }
        return chart.Title.Length <= 32_767 && !HasControls(chart.Title);
    }

    private static bool TryReadComboSeries(XElement plot, SpreadsheetChartType type, PresentationChartAxisGroup axisGroup, out ComboNativeSeries[] result)
    {
        result = [];
        var nativeSeries = plot.Elements(ChartNs + "ser").ToArray();
        if (nativeSeries.Length == 0) return false;
        var output = new List<ComboNativeSeries>();
        foreach (var series in nativeSeries)
        {
            if (!TryComboSeriesOrder(series, "idx", out var index) || !TryComboSeriesOrder(series, "order", out var order) || index != order) return false;
            output.Add(new ComboNativeSeries(type, axisGroup, series, order));
        }
        result = output.ToArray();
        return true;
    }

    private static bool TryComboSeriesOrder(XElement series, string name, out uint value)
    {
        value = 0;
        var elements = series.Elements(ChartNs + name).Take(2).ToArray();
        if (elements.Length != 1 || elements[0].Elements().Any() || elements[0].Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value))) return false;
        var attributes = elements[0].Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToArray();
        return attributes.Length == 1 && attributes[0].Name == "val" && uint.TryParse(attributes[0].Value, NumberStyles.None, CultureInfo.InvariantCulture, out value);
    }

    private static bool IsCanonicalComboPlot(XElement plot, SpreadsheetChartType type)
    {
        var allowed = type == SpreadsheetChartType.Bar
            ? new HashSet<XName> { ChartNs + "barDir", ChartNs + "grouping", ChartNs + "ser", ChartNs + "dLbls", ChartNs + "axId" }
            : new HashSet<XName> { ChartNs + "grouping", ChartNs + "ser", ChartNs + "dLbls", ChartNs + "axId" };
        if (plot.Elements().Any(item => !allowed.Contains(item.Name))) return false;
        if (!ComboScalarEquals(plot.Element(ChartNs + "grouping"), type == SpreadsheetChartType.Bar ? "clustered" : "standard")) return false;
        if (type == SpreadsheetChartType.Bar && !ComboScalarEquals(plot.Element(ChartNs + "barDir"), "col")) return false;
        var axisIds = plot.Elements(ChartNs + "axId").ToArray();
        return axisIds.Length == 2 && axisIds.All(item => ComboScalar(item, out _));
    }

    private static bool SharesPrimaryAxes(XElement barPlot, XElement linePlot)
    {
        var barIds = barPlot.Elements(ChartNs + "axId").Select(item => ComboScalar(item, out var value) ? value : string.Empty).ToArray();
        var lineIds = linePlot.Elements(ChartNs + "axId").Select(item => ComboScalar(item, out var value) ? value : string.Empty).ToArray();
        return barIds.Length == 2 && lineIds.Length == 2 && barIds.Distinct(StringComparer.Ordinal).Count() == 2 && barIds.SequenceEqual(lineIds, StringComparer.Ordinal);
    }

    private static bool ComboScalarEquals(XElement? element, string expected) => ComboScalar(element, out var value) && value == expected;

    private static bool ComboScalar(XElement? element, out string value)
    {
        value = string.Empty;
        if (element is null || element.Elements().Any() || element.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value))) return false;
        var attributes = element.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToArray();
        if (attributes.Length != 1 || attributes[0].Name != "val") return false;
        value = attributes[0].Value;
        return value.Length > 0;
    }

    private static XDocument BuildComboChartDocument(PresentationChart chart, string id, string name)
    {
        var indexed = chart.ComboSeries.Select((item, index) => new ComboNativeSeries(item.Type, ComboAxisGroup(item), SeriesElement(item.Series, chart.Categories, index), checked((uint)index))).ToArray();
        var barSeries = indexed.Where(item => item.Type == SpreadsheetChartType.Bar).Select(item => item.Element).ToArray();
        var lineSeries = indexed.Where(item => item.Type == SpreadsheetChartType.Line).Select(item => item.Element).ToArray();
        var lineUsesSecondaryAxes = indexed.Any(item => item.Type == SpreadsheetChartType.Line && item.AxisGroup == PresentationChartAxisGroup.Secondary);
        var barPlot = new XElement(ChartNs + "barChart",
            new XElement(ChartNs + "barDir", new XAttribute("val", "col")),
            new XElement(ChartNs + "grouping", new XAttribute("val", "clustered")),
            barSeries,
            XlsxChartDataLabelsCodec.Element(chart.DataLabels),
            new XElement(ChartNs + "axId", new XAttribute("val", "1")),
            new XElement(ChartNs + "axId", new XAttribute("val", "2")));
        var linePlot = new XElement(ChartNs + "lineChart",
            new XElement(ChartNs + "grouping", new XAttribute("val", "standard")),
            lineSeries,
            XlsxChartDataLabelsCodec.Element(chart.DataLabels),
            new XElement(ChartNs + "axId", new XAttribute("val", lineUsesSecondaryAxes ? "3" : "1")),
            new XElement(ChartNs + "axId", new XAttribute("val", lineUsesSecondaryAxes ? "4" : "2")));
        var plotArea = new XElement(ChartNs + "plotArea", new XElement(ChartNs + "layout"), barPlot, linePlot);
        XlsxChartAxisCodec.AppendAuthored(plotArea, ComboAxisCarrier(chart, id, name));
        if (lineUsesSecondaryAxes) XlsxChartAxisCodec.AppendAuthoredPresentationSecondary(plotArea, ComboAxisCarrier(chart, id, name, secondary: true));
        var nativeChart = new XElement(ChartNs + "chart");
        if (chart.Title.Length > 0) nativeChart.Add(XlsxChartTextStyleCodec.TitleElement(chart.Title, null));
        nativeChart.Add(plotArea);
        if (chart.HasLegend) nativeChart.Add(LegendElement());
        nativeChart.Add(new XElement(ChartNs + "plotVisOnly", new XAttribute("val", "1")));
        return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), new XElement(ChartNs + "chartSpace", new XAttribute(XNamespace.Xmlns + "c", ChartNs), new XAttribute(XNamespace.Xmlns + "a", DrawingNs), nativeChart));
    }

    private static void PatchComboChart(XDocument document, PresentationChart target, string id, string name)
    {
        var nativeChart = document.Root!.Element(ChartNs + "chart")!;
        PatchTitle(nativeChart, target.Title);
        PatchLegend(nativeChart, target.HasLegend);
        var plotArea = nativeChart.Element(ChartNs + "plotArea")!;
        var barPlot = plotArea.Element(ChartNs + "barChart")!;
        var linePlot = plotArea.Element(ChartNs + "lineChart")!;
        var lineUsesSecondaryAxes = !SharesPrimaryAxes(barPlot, linePlot);
        if (lineUsesSecondaryAxes != HasSecondaryComboLine(target) ||
            !TryReadComboSeries(barPlot, SpreadsheetChartType.Bar, PresentationChartAxisGroup.Primary, out var barSeries) ||
            !TryReadComboSeries(linePlot, SpreadsheetChartType.Line, lineUsesSecondaryAxes ? PresentationChartAxisGroup.Secondary : PresentationChartAxisGroup.Primary, out var lineSeries))
            throw new CodecException("unsupported_presentation_edit", "Presentation combo chart no longer matches the bounded native series profile.");
        var nativeSeries = barSeries.Concat(lineSeries).OrderBy(item => item.Order).ToArray();
        if (nativeSeries.Length != target.ComboSeries.Count || nativeSeries.Select(item => item.Order).Distinct().Count() != nativeSeries.Length ||
            !nativeSeries.Select(item => item.Order).SequenceEqual(Enumerable.Range(0, nativeSeries.Length).Select(index => (uint)index)))
            throw new CodecException("presentation_chart_topology_changed", "Presentation combo chart series topology changed unexpectedly.");
        for (var index = 0; index < nativeSeries.Length; index++)
        {
            var requested = target.ComboSeries[index];
            if (requested.Type != nativeSeries[index].Type || ComboAxisGroup(requested) != nativeSeries[index].AxisGroup || requested.Series is null) throw new CodecException("presentation_chart_topology_changed", "Presentation combo chart series type or axis group changed unexpectedly.");
            PatchSeries(nativeSeries[index].Element, requested.Series, target.Categories);
        }
        XlsxChartDataLabelsCodec.Patch(barPlot, target.DataLabels);
        XlsxChartDataLabelsCodec.Patch(linePlot, target.DataLabels);
        XlsxChartAxisCodec.Patch(plotArea, barPlot, ComboAxisCarrier(target, id, name));
        if (lineUsesSecondaryAxes) XlsxChartAxisCodec.PatchPresentationSecondary(plotArea, linePlot, ComboAxisCarrier(target, id, name, secondary: true));
    }
}
