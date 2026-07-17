using System.Globalization;
using System.Xml.Linq;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns the bounded primary category/value or scatter value/value-axis projection for worksheet charts.
// Axis identity and all unmodeled formatting remain in the ChartPart; this
// module reads and patches only titles, number formats, category label interval,
// linear value-axis bounds/unit, and the delegated bounded tick-label style.
internal static class XlsxChartAxisCodec
{
    private const uint MaxTickLabelInterval = 1_048_576;
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    internal static void Validate(SpreadsheetChartArtifact chart, string worksheetId)
    {
        if (chart.Type is SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut)
        {
            if (chart.XAxis is not null || chart.YAxis is not null)
                throw Invalid(worksheetId, chart.Id, "pie and doughnut charts cannot carry category/value axes in the bounded profile.");
            return;
        }
        if ((chart.XAxis is null) != (chart.YAxis is null))
            throw Invalid(worksheetId, chart.Id, "must carry both x_axis and y_axis or neither for backward-compatible default authoring.");
        if (chart.XAxis is null) return;
        ValidateAxis(chart.XAxis, chart.Type != SpreadsheetChartType.Scatter, "x", worksheetId, chart.Id);
        ValidateAxis(chart.YAxis!, false, "y", worksheetId, chart.Id);
    }

    internal static bool TryRead(XElement plotArea, XElement plot, SpreadsheetChartArtifact chart, out bool editable)
    {
        editable = false;
        if (chart.Type is SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut)
        {
            editable = !plotArea.Elements().Any(IsAxis);
            return editable;
        }
        var scatter = chart.Type == SpreadsheetChartType.Scatter;
        if (!TryLocate(plotArea, plot, scatter, out var horizontalAxis, out var verticalAxis)) return false;
        if (!TryReadAxis(horizontalAxis, !scatter, "b", out var xAxis, out var xEditable) ||
            !TryReadAxis(verticalAxis, false, "l", out var yAxis, out var yEditable)) return false;
        chart.XAxis = xAxis;
        chart.YAxis = yAxis;
        editable = xEditable && yEditable;
        return true;
    }

    internal static void AppendAuthored(XElement plotArea, SpreadsheetChartArtifact chart)
    {
        if (chart.Type is SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut) return;
        var xAxis = chart.XAxis ?? new SpreadsheetChartAxisArtifact();
        var yAxis = chart.YAxis ?? new SpreadsheetChartAxisArtifact();
        if (chart.Type == SpreadsheetChartType.Scatter)
            plotArea.Add(BuildValueAxis(xAxis, "1", "2", "b"), BuildValueAxis(yAxis, "2", "1", "l"));
        else
            plotArea.Add(BuildCategoryAxis(xAxis), BuildValueAxis(yAxis, "2", "1", "l"));
    }

    internal static void Patch(XElement plotArea, XElement plot, SpreadsheetChartArtifact target)
    {
        if (target.Type is SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut) return;
        var scatter = target.Type == SpreadsheetChartType.Scatter;
        if (target.XAxis is null || target.YAxis is null || !TryLocate(plotArea, plot, scatter, out var horizontalAxis, out var verticalAxis))
            throw new CodecException("unsupported_spreadsheet_chart_edit", $"Worksheet chart {target.Id} cannot change its primary-axis topology.");
        PatchAxis(horizontalAxis, target.XAxis, !scatter);
        PatchAxis(verticalAxis, target.YAxis, false);
    }

    internal static string Semantics(SpreadsheetChartArtifact chart) =>
        string.Join('\u001d', AxisSemantics(chart.XAxis), AxisSemantics(chart.YAxis));

    private static void ValidateAxis(SpreadsheetChartAxisArtifact axis, bool category, string axisName, string worksheetId, string chartId)
    {
        if (axis.Title.Length > 32_767 || HasControls(axis.Title)) throw Invalid(worksheetId, chartId, $"{axisName}-axis title is invalid.");
        if (axis.NumberFormatCode.Length > 255 || HasControls(axis.NumberFormatCode)) throw Invalid(worksheetId, chartId, $"{axisName}-axis number format is invalid.");
        if (category)
        {
            if (axis.HasMinimum || axis.HasMaximum || axis.HasMajorUnit) throw Invalid(worksheetId, chartId, $"{axisName}-axis cannot carry numeric minimum, maximum, or major unit.");
            if (axis.HasTickLabelInterval && axis.TickLabelInterval is < 1 or > MaxTickLabelInterval) throw Invalid(worksheetId, chartId, $"{axisName}-axis tick label interval must be 1 through {MaxTickLabelInterval}.");
            return;
        }
        if (axis.HasTickLabelInterval) throw Invalid(worksheetId, chartId, $"{axisName}-axis cannot carry a category tick label interval.");
        if (axis.HasMinimum && !double.IsFinite(axis.Minimum) || axis.HasMaximum && !double.IsFinite(axis.Maximum)) throw Invalid(worksheetId, chartId, $"{axisName}-axis minimum and maximum must be finite.");
        if (axis.HasMinimum && axis.HasMaximum && axis.Minimum >= axis.Maximum) throw Invalid(worksheetId, chartId, $"{axisName}-axis minimum must be less than maximum.");
        if (axis.HasMajorUnit && (!double.IsFinite(axis.MajorUnit) || axis.MajorUnit <= 0)) throw Invalid(worksheetId, chartId, $"{axisName}-axis major unit must be finite and positive.");
    }

    private static bool TryLocate(XElement plotArea, XElement plot, bool scatter, out XElement horizontalAxis, out XElement verticalAxis)
    {
        horizontalAxis = null!;
        verticalAxis = null!;
        var axes = plotArea.Elements().Where(IsAxis).ToArray();
        var categories = axes.Where(item => item.Name == ChartNs + "catAx").ToArray();
        var values = axes.Where(item => item.Name == ChartNs + "valAx").ToArray();
        if (axes.Length != 2) return false;
        if (scatter)
        {
            if (categories.Length != 0 || values.Length != 2) return false;
            var horizontal = values.Where(item => AxisValue(item.Element(ChartNs + "axPos")) == "b").ToArray();
            var vertical = values.Where(item => AxisValue(item.Element(ChartNs + "axPos")) == "l").ToArray();
            if (horizontal.Length != 1 || vertical.Length != 1 || ReferenceEquals(horizontal[0], vertical[0])) return false;
            horizontalAxis = horizontal[0];
            verticalAxis = vertical[0];
        }
        else
        {
            if (categories.Length != 1 || values.Length != 1) return false;
            horizontalAxis = categories[0];
            verticalAxis = values[0];
        }
        var plotIds = plot.Elements(ChartNs + "axId").Select(AxisValue).ToArray();
        var horizontalId = AxisValue(horizontalAxis.Element(ChartNs + "axId"));
        var verticalId = AxisValue(verticalAxis.Element(ChartNs + "axId"));
        if (plotIds.Length != 2 || plotIds.Any(string.IsNullOrEmpty) || plotIds.Distinct(StringComparer.Ordinal).Count() != 2 ||
            string.IsNullOrEmpty(horizontalId) || string.IsNullOrEmpty(verticalId) || horizontalId == verticalId ||
            !plotIds.Contains(horizontalId, StringComparer.Ordinal) || !plotIds.Contains(verticalId, StringComparer.Ordinal)) return false;
        return AxisValue(horizontalAxis.Element(ChartNs + "crossAx")) == verticalId && AxisValue(verticalAxis.Element(ChartNs + "crossAx")) == horizontalId;
    }

    private static bool TryReadAxis(XElement source, bool category, string expectedPosition, out SpreadsheetChartAxisArtifact axis, out bool editable)
    {
        axis = new SpreadsheetChartAxisArtifact();
        editable = true;
        if (!Singleton(source, "scaling", out var scaling) || scaling is null ||
            !Singleton(scaling, "orientation", out var orientation) || orientation is null || AxisValue(orientation) != "minMax") return false;
        if (scaling.Element(ChartNs + "logBase") is not null) editable = false;
        if (!Singleton(source, "delete", out var deleted)) return false;
        if (deleted is not null && !IsFalse(AxisValue(deleted))) editable = false;
        if (!Singleton(source, "axPos", out var position) || position is null) return false;
        if (AxisValue(position) != expectedPosition) editable = false;
        if (!TryTitle(source, out var title, out var titleEditable) || !TryNumberFormat(source, out var numberFormat, out var numberFormatEditable)) return false;
        axis.Title = title;
        axis.NumberFormatCode = numberFormat;
        editable &= titleEditable && numberFormatEditable && XlsxChartTextStyleCodec.TryReadAxis(source, axis);
        if (category)
        {
            if (scaling.Element(ChartNs + "min") is not null || scaling.Element(ChartNs + "max") is not null) editable = false;
            if (!TryOptionalUInt(source, "tickLblSkip", out var hasInterval, out var interval) || hasInterval && interval is < 1 or > MaxTickLabelInterval) return false;
            if (hasInterval) axis.TickLabelInterval = interval;
            return true;
        }
        if (!TryOptionalDouble(scaling, "min", out var hasMinimum, out var minimum) ||
            !TryOptionalDouble(scaling, "max", out var hasMaximum, out var maximum) ||
            !TryOptionalDouble(source, "majorUnit", out var hasMajorUnit, out var majorUnit)) return false;
        if (hasMinimum) axis.Minimum = minimum;
        if (hasMaximum) axis.Maximum = maximum;
        if (hasMajorUnit) axis.MajorUnit = majorUnit;
        if (hasMinimum && hasMaximum && minimum >= maximum || hasMajorUnit && majorUnit <= 0) return false;
        return true;
    }

    private static bool TryTitle(XElement source, out string title, out bool editable)
    {
        title = string.Empty;
        editable = true;
        if (!Singleton(source, "title", out var element)) return false;
        if (element is null) return true;
        var richText = element.Descendants(DrawingNs + "t").ToArray();
        if (richText.Length > 0) title = string.Concat(richText.Select(item => item.Value));
        else
        {
            title = element.Descendants(ChartNs + "v").FirstOrDefault()?.Value ?? string.Empty;
            editable = false;
        }
        return title.Length <= 32_767 && !HasControls(title);
    }

    private static bool TryNumberFormat(XElement source, out string code, out bool editable)
    {
        code = string.Empty;
        editable = true;
        if (!Singleton(source, "numFmt", out var element)) return false;
        if (element is null) return true;
        code = (string?)element.Attribute("formatCode") ?? string.Empty;
        if (code.Length > 255 || HasControls(code) || element.Attribute("formatCode") is null) return false;
        var sourceLinked = (string?)element.Attribute("sourceLinked");
        if (sourceLinked is null || !IsFalse(sourceLinked)) editable = false;
        return true;
    }

    private static XElement BuildCategoryAxis(SpreadsheetChartAxisArtifact axis)
    {
        var output = new XElement(ChartNs + "catAx",
            new XElement(ChartNs + "axId", new XAttribute("val", "1")),
            new XElement(ChartNs + "scaling", new XElement(ChartNs + "orientation", new XAttribute("val", "minMax"))),
            new XElement(ChartNs + "axPos", new XAttribute("val", "b")));
        AppendTitleAndNumberFormat(output, axis);
        XlsxChartTextStyleCodec.AppendAuthoredAxis(output, axis.TextStyle);
        output.Add(new XElement(ChartNs + "crossAx", new XAttribute("val", "2")));
        if (axis.HasTickLabelInterval) output.Add(ValueElement("tickLblSkip", axis.TickLabelInterval));
        return output;
    }

    private static XElement BuildValueAxis(SpreadsheetChartAxisArtifact axis, string axisId, string crossAxisId, string position)
    {
        var scaling = new XElement(ChartNs + "scaling", new XElement(ChartNs + "orientation", new XAttribute("val", "minMax")));
        if (axis.HasMaximum) scaling.Add(ValueElement("max", axis.Maximum));
        if (axis.HasMinimum) scaling.Add(ValueElement("min", axis.Minimum));
        var output = new XElement(ChartNs + "valAx",
            new XElement(ChartNs + "axId", new XAttribute("val", axisId)), scaling,
            new XElement(ChartNs + "axPos", new XAttribute("val", position)));
        AppendTitleAndNumberFormat(output, axis);
        XlsxChartTextStyleCodec.AppendAuthoredAxis(output, axis.TextStyle);
        output.Add(new XElement(ChartNs + "crossAx", new XAttribute("val", crossAxisId)));
        if (axis.HasMajorUnit) output.Add(ValueElement("majorUnit", axis.MajorUnit));
        return output;
    }

    private static void AppendTitleAndNumberFormat(XElement axis, SpreadsheetChartAxisArtifact semantic)
    {
        if (semantic.Title.Length > 0) axis.Add(TitleElement(semantic.Title));
        if (semantic.NumberFormatCode.Length > 0) axis.Add(NumberFormatElement(semantic.NumberFormatCode));
    }

    private static void PatchAxis(XElement native, SpreadsheetChartAxisArtifact target, bool category)
    {
        PatchTitle(native, target.Title);
        PatchNumberFormat(native, target.NumberFormatCode);
        XlsxChartTextStyleCodec.PatchAxis(native, target.TextStyle);
        if (category)
        {
            PatchValue(native, "tickLblSkip", target.HasTickLabelInterval, target.TickLabelInterval, ["tickMarkSkip", "noMultiLvlLbl", "extLst"]);
            return;
        }
        var scaling = native.Element(ChartNs + "scaling")!;
        PatchValue(scaling, "max", target.HasMaximum, target.Maximum, ["min", "extLst"]);
        PatchValue(scaling, "min", target.HasMinimum, target.Minimum, ["extLst"]);
        PatchValue(native, "majorUnit", target.HasMajorUnit, target.MajorUnit, ["minorUnit", "dispUnits", "extLst"]);
    }

    private static void PatchTitle(XElement owner, string title)
    {
        var existing = owner.Element(ChartNs + "title");
        if (title.Length == 0) { existing?.Remove(); return; }
        if (existing is null)
        {
            InsertBefore(owner, TitleElement(title), ["numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt", "extLst"]);
            return;
        }
        var runs = existing.Descendants(DrawingNs + "t").ToArray();
        if (runs.Length == 0) throw new CodecException("unsupported_spreadsheet_chart_edit", "Referenced worksheet-chart axis titles are read-only.");
        runs[0].Value = title;
        foreach (var run in runs.Skip(1)) run.Value = string.Empty;
    }

    private static void PatchNumberFormat(XElement owner, string code)
    {
        var existing = owner.Element(ChartNs + "numFmt");
        if (code.Length == 0) { existing?.Remove(); return; }
        if (existing is null)
        {
            InsertBefore(owner, NumberFormatElement(code), ["majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt", "extLst"]);
            return;
        }
        existing.SetAttributeValue("formatCode", code);
        existing.SetAttributeValue("sourceLinked", "0");
    }

    private static void PatchValue<T>(XElement owner, string name, bool present, T value, string[] laterNames)
    {
        var existing = owner.Element(ChartNs + name);
        if (!present) { existing?.Remove(); return; }
        if (existing is null) InsertBefore(owner, ValueElement(name, value), laterNames);
        else existing.SetAttributeValue("val", Format(value));
    }

    private static void InsertBefore(XElement owner, XElement value, IEnumerable<string> laterNames)
    {
        var later = new HashSet<XName>(laterNames.Select(name => ChartNs + name));
        var next = owner.Elements().FirstOrDefault(item => later.Contains(item.Name));
        if (next is null) owner.Add(value);
        else next.AddBeforeSelf(value);
    }

    private static XElement TitleElement(string title) => new(ChartNs + "title",
        new XElement(ChartNs + "tx", new XElement(ChartNs + "rich",
            new XElement(DrawingNs + "bodyPr"), new XElement(DrawingNs + "lstStyle"),
            new XElement(DrawingNs + "p", new XElement(DrawingNs + "r", new XElement(DrawingNs + "t", title))))),
        new XElement(ChartNs + "layout"));

    private static XElement NumberFormatElement(string code) => new(ChartNs + "numFmt", new XAttribute("formatCode", code), new XAttribute("sourceLinked", "0"));
    private static XElement ValueElement<T>(string name, T value) => new(ChartNs + name, new XAttribute("val", Format(value)));
    private static string Format<T>(T value) => value switch { double number => number.ToString("R", CultureInfo.InvariantCulture), IFormattable item => item.ToString(null, CultureInfo.InvariantCulture), _ => value?.ToString() ?? string.Empty };

    private static bool TryOptionalUInt(XElement source, string name, out bool present, out uint value)
    {
        present = false; value = 0;
        if (!Singleton(source, name, out var element)) return false;
        if (element is null) return true;
        present = true;
        return uint.TryParse(AxisValue(element), NumberStyles.None, CultureInfo.InvariantCulture, out value);
    }

    private static bool TryOptionalDouble(XElement source, string name, out bool present, out double value)
    {
        present = false; value = 0;
        if (!Singleton(source, name, out var element)) return false;
        if (element is null) return true;
        present = true;
        return double.TryParse(AxisValue(element), NumberStyles.Float, CultureInfo.InvariantCulture, out value) && double.IsFinite(value);
    }

    private static bool Singleton(XElement source, string name, out XElement? element)
    {
        var matches = source.Elements(ChartNs + name).Take(2).ToArray();
        element = matches.FirstOrDefault();
        return matches.Length <= 1;
    }

    private static string AxisValue(XElement? source) => (string?)source?.Attribute("val") ?? string.Empty;
    private static bool IsAxis(XElement source) => source.Name == ChartNs + "catAx" || source.Name == ChartNs + "dateAx" || source.Name == ChartNs + "valAx" || source.Name == ChartNs + "serAx";
    private static bool IsFalse(string value) => value is "0" or "false" or "off";
    private static bool HasControls(string value) => value.Any(char.IsControl);
    private static string AxisSemantics(SpreadsheetChartAxisArtifact? axis) => axis is null ? "-" : string.Join('\u001f', axis.Title, axis.NumberFormatCode,
        axis.HasTickLabelInterval ? axis.TickLabelInterval.ToString(CultureInfo.InvariantCulture) : "-",
        axis.HasMinimum ? axis.Minimum.ToString("R", CultureInfo.InvariantCulture) : "-",
        axis.HasMaximum ? axis.Maximum.ToString("R", CultureInfo.InvariantCulture) : "-",
        axis.HasMajorUnit ? axis.MajorUnit.ToString("R", CultureInfo.InvariantCulture) : "-",
        XlsxChartTextStyleCodec.Semantics(axis.TextStyle));
    private static CodecException Invalid(string worksheetId, string chartId, string message) => new("invalid_spreadsheet_chart", $"Worksheet {worksheetId} chart {chartId} {message}");
}
