using System.Xml.Linq;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns one chart-level c:smooth scalar on a line plot. Grouping, color
// variation, and all other line-plot options remain separate future slices.
internal static class XlsxChartLineOptionsCodec
{
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    internal static void Validate(SpreadsheetChartArtifact chart, string worksheetId)
    {
        if (chart.LineOptions is null) return;
        if (chart.Type != SpreadsheetChartType.Line)
            throw new CodecException("invalid_spreadsheet_chart", $"Worksheet {worksheetId} chart {chart.Id} line options are supported only on line charts.");
        if (!chart.LineOptions.HasSmooth)
            throw new CodecException("invalid_spreadsheet_chart", $"Worksheet {worksheetId} chart {chart.Id} line options must carry explicit smooth presence.");
    }

    internal static bool TryRead(XElement plot, SpreadsheetChartArtifact chart)
    {
        var elements = plot.Elements(ChartNs + "smooth").ToArray();
        if (elements.Length == 0) return true;
        if (elements.Length != 1 || chart.Type != SpreadsheetChartType.Line) return false;
        var native = elements[0];
        if (native.Nodes().Any(node => node is XText text ? !string.IsNullOrWhiteSpace(text.Value) : true) || native.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != "val")) return false;
        var value = (string?)native.Attribute("val");
        if (value is not ("0" or "1" or "false" or "true")) return false;
        chart.LineOptions = new SpreadsheetChartLineOptionsArtifact { Smooth = value is "1" or "true" };
        return true;
    }

    internal static XElement? Element(SpreadsheetChartLineOptionsArtifact? options) => options?.HasSmooth == true
        ? new XElement(ChartNs + "smooth", new XAttribute("val", options.Smooth ? "1" : "0"))
        : null;

    internal static void Patch(XElement plot, SpreadsheetChartLineOptionsArtifact? options)
    {
        var existing = plot.Element(ChartNs + "smooth");
        var replacement = Element(options);
        if (replacement is null) { existing?.Remove(); return; }
        if (existing is not null) { existing.ReplaceWith(replacement); return; }
        var axis = plot.Elements(ChartNs + "axId").FirstOrDefault();
        if (axis is null) plot.Add(replacement);
        else axis.AddBeforeSelf(replacement);
    }

    internal static string Semantics(SpreadsheetChartLineOptionsArtifact? options) => options is null
        ? "no-line-options"
        : options.HasSmooth ? $"smooth:{(options.Smooth ? 1 : 0)}" : "line-options-without-smooth";
}
