using System.Xml.Linq;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns the required c:grouping scalar plus optional c:varyColors and c:smooth
// scalars on one line plot. Every other line-plot option remains separate.
internal static class XlsxChartLineOptionsCodec
{
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly HashSet<string> GroupingValues = new(StringComparer.Ordinal) { "standard", "stacked", "percentStacked" };
    private static readonly HashSet<string> BooleanValues = new(StringComparer.Ordinal) { "0", "1", "false", "true" };

    internal static void Validate(SpreadsheetChartArtifact chart, string worksheetId)
    {
        if (chart.LineOptions is null) return;
        if (chart.Type != SpreadsheetChartType.Line)
            throw new CodecException("invalid_spreadsheet_chart", $"Worksheet {worksheetId} chart {chart.Id} line options are supported only on line charts.");
        if (!chart.LineOptions.HasGrouping && !chart.LineOptions.HasSmooth && !chart.LineOptions.VaryColors)
            throw new CodecException("invalid_spreadsheet_chart", $"Worksheet {worksheetId} chart {chart.Id} line options must carry explicit grouping, smooth, or vary-colors semantics.");
        if (chart.LineOptions.HasGrouping && chart.LineOptions.Grouping is not (SpreadsheetChartLineGrouping.Standard or SpreadsheetChartLineGrouping.Stacked or SpreadsheetChartLineGrouping.PercentStacked))
            throw new CodecException("invalid_spreadsheet_chart", $"Worksheet {worksheetId} chart {chart.Id} line grouping is unsupported.");
    }

    internal static bool TryRead(XElement plot, SpreadsheetChartArtifact chart)
    {
        var smoothElements = plot.Elements(ChartNs + "smooth").ToArray();
        if (chart.Type != SpreadsheetChartType.Line) return smoothElements.Length == 0;

        var options = new SpreadsheetChartLineOptionsArtifact();
        var editable = TryReadGrouping(plot, options);
        editable &= TryReadVaryColors(plot, options);
        if (smoothElements.Length == 1)
        {
            var native = smoothElements[0];
            if (IsScalar(native, BooleanValues, out var value)) options.Smooth = value is "1" or "true";
            else editable = false;
        }
        else if (smoothElements.Length > 1) editable = false;

        if (options.HasGrouping || options.HasSmooth) chart.LineOptions = options;
        return editable;
    }

    internal static XElement GroupingElement(SpreadsheetChartLineOptionsArtifact? options) =>
        CreateGroupingElement(EffectiveGrouping(options));

    internal static XElement? SmoothElement(SpreadsheetChartLineOptionsArtifact? options) => options?.HasSmooth == true
        ? CreateSmoothElement(options.Smooth)
        : null;

    internal static XElement? VaryColorsElement(SpreadsheetChartLineOptionsArtifact? options) => options?.VaryColors == true
        ? CreateVaryColorsElement()
        : null;

    internal static void Patch(XElement plot, SpreadsheetChartLineOptionsArtifact? options)
    {
        var grouping = CreateGroupingElement(EffectiveGrouping(options));
        var existingGrouping = plot.Element(ChartNs + "grouping");
        if (existingGrouping is not null) existingGrouping.ReplaceWith(grouping);
        else
        {
            var series = plot.Element(ChartNs + "ser");
            if (series is null) plot.AddFirst(grouping);
            else series.AddBeforeSelf(grouping);
        }

        var existingVaryColors = plot.Element(ChartNs + "varyColors");
        var replacementVaryColors = options?.VaryColors == true ? CreateVaryColorsElement() : null;
        if (replacementVaryColors is null) existingVaryColors?.Remove();
        else if (existingVaryColors is not null) existingVaryColors.ReplaceWith(replacementVaryColors);
        else
        {
            var series = plot.Element(ChartNs + "ser");
            if (series is null) grouping.AddAfterSelf(replacementVaryColors);
            else series.AddBeforeSelf(replacementVaryColors);
        }

        var existingSmooth = plot.Element(ChartNs + "smooth");
        var replacementSmooth = options?.HasSmooth == true ? CreateSmoothElement(options.Smooth) : null;
        if (replacementSmooth is null) { existingSmooth?.Remove(); return; }
        if (existingSmooth is not null) { existingSmooth.ReplaceWith(replacementSmooth); return; }
        var axis = plot.Elements(ChartNs + "axId").FirstOrDefault();
        if (axis is null) plot.Add(replacementSmooth);
        else axis.AddBeforeSelf(replacementSmooth);
    }

    internal static string Semantics(SpreadsheetChartLineOptionsArtifact? options)
    {
        var grouping = GroupingValue(EffectiveGrouping(options));
        var smooth = options?.HasSmooth == true ? (options.Smooth ? "1" : "0") : "absent";
        return $"grouping:{grouping};vary-colors:{(options?.VaryColors == true ? 1 : 0)};smooth:{smooth}";
    }

    private static bool TryReadGrouping(XElement plot, SpreadsheetChartLineOptionsArtifact options)
    {
        var elements = plot.Elements(ChartNs + "grouping").ToArray();
        if (elements.Length != 1 || !IsScalar(elements[0], GroupingValues, out var value)) return false;
        options.Grouping = value switch
        {
            "standard" => SpreadsheetChartLineGrouping.Standard,
            "stacked" => SpreadsheetChartLineGrouping.Stacked,
            "percentStacked" => SpreadsheetChartLineGrouping.PercentStacked,
            _ => throw new InvalidOperationException(),
        };
        return true;
    }

    private static bool TryReadVaryColors(XElement plot, SpreadsheetChartLineOptionsArtifact options)
    {
        var elements = plot.Elements(ChartNs + "varyColors").ToArray();
        if (elements.Length == 0) return true;
        if (elements.Length != 1 || !IsScalar(elements[0], BooleanValues, out var value)) return false;
        options.VaryColors = value is "1" or "true";
        return true;
    }

    private static bool IsScalar(XElement element, IReadOnlySet<string> allowed, out string? value)
    {
        value = (string?)element.Attribute("val");
        return !element.Nodes().Any(node => node is XText text ? !string.IsNullOrWhiteSpace(text.Value) : true)
            && !element.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != "val")
            && value is not null
            && allowed.Contains(value);
    }

    private static SpreadsheetChartLineGrouping EffectiveGrouping(SpreadsheetChartLineOptionsArtifact? options) => options?.HasGrouping == true
        ? options.Grouping
        : SpreadsheetChartLineGrouping.Standard;

    private static XElement CreateGroupingElement(SpreadsheetChartLineGrouping grouping) =>
        new(ChartNs + "grouping", new XAttribute("val", GroupingValue(grouping)));

    private static XElement CreateSmoothElement(bool value) =>
        new(ChartNs + "smooth", new XAttribute("val", value ? "1" : "0"));

    private static XElement CreateVaryColorsElement() =>
        new(ChartNs + "varyColors", new XAttribute("val", "1"));

    private static string GroupingValue(SpreadsheetChartLineGrouping grouping) => grouping switch
    {
        SpreadsheetChartLineGrouping.Standard => "standard",
        SpreadsheetChartLineGrouping.Stacked => "stacked",
        SpreadsheetChartLineGrouping.PercentStacked => "percentStacked",
        _ => throw new InvalidOperationException("Validated line grouping is unsupported."),
    };
}
