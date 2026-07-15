using System.Xml.Linq;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns one plot-level c:dLbls container with direct c:showVal and
// c:showCatName booleans. Standard unsupported show flags are accepted only
// when false and are retained byte-for-byte during another bounded edit.
internal static class XlsxChartDataLabelsCodec
{
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly string[] OrderedFlags = ["showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize"];
    private static readonly HashSet<string> AllowedFlags = new(OrderedFlags, StringComparer.Ordinal);
    private static readonly HashSet<string> BooleanValues = new(StringComparer.Ordinal) { "0", "1", "false", "true" };

    internal static void Validate(SpreadsheetChartArtifact chart, string worksheetId)
    {
        // Both fields are ordinary protobuf booleans. Message presence is the
        // only additional state and is valid even when both values are false.
        _ = chart.DataLabels;
        _ = worksheetId;
    }

    internal static bool TryRead(XElement plot, SpreadsheetChartArtifact chart)
    {
        var containers = plot.Elements(ChartNs + "dLbls").ToArray();
        if (containers.Length == 0) return true;
        if (containers.Length != 1) return false;
        var labels = containers[0];
        if (labels.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) || HasUnexpectedText(labels)) return false;
        var children = labels.Elements().ToArray();
        if (children.Any(child => child.Name.Namespace != ChartNs || !AllowedFlags.Contains(child.Name.LocalName)) ||
            AllowedFlags.Any(name => children.Count(child => child.Name == ChartNs + name) > 1)) return false;
        var showValue = children.SingleOrDefault(child => child.Name == ChartNs + "showVal");
        var showCategoryName = children.SingleOrDefault(child => child.Name == ChartNs + "showCatName");
        if (!TryBoolean(showValue, out var value) || !TryBoolean(showCategoryName, out var categoryName)) return false;
        foreach (var name in OrderedFlags.Where(name => name is not "showVal" and not "showCatName"))
        {
            var element = children.SingleOrDefault(child => child.Name == ChartNs + name);
            if (element is not null && (!TryBoolean(element, out var enabled) || enabled)) return false;
        }
        chart.DataLabels = new SpreadsheetChartDataLabelsArtifact { ShowValue = value, ShowCategoryName = categoryName };
        return true;
    }

    internal static XElement? Element(SpreadsheetChartDataLabelsArtifact? labels) => labels is null ? null :
        new XElement(ChartNs + "dLbls",
            BooleanElement("showVal", labels.ShowValue),
            BooleanElement("showCatName", labels.ShowCategoryName));

    internal static void Patch(XElement plot, SpreadsheetChartDataLabelsArtifact? labels)
    {
        var existing = plot.Element(ChartNs + "dLbls");
        if (labels is null) { existing?.Remove(); return; }
        if (existing is null)
        {
            var created = Element(labels)!;
            var lastSeries = plot.Elements(ChartNs + "ser").LastOrDefault();
            if (lastSeries is not null) lastSeries.AddAfterSelf(created);
            else
            {
                var firstAfterSeries = plot.Elements().FirstOrDefault(element => element.Name.LocalName is "dLbls" or "dropLines" or "hiLowLines" or "upDownBars" or "marker" or "smooth" or "gapWidth" or "overlap" or "firstSliceAng" or "holeSize" or "axId" or "extLst");
                if (firstAfterSeries is null) plot.Add(created);
                else firstAfterSeries.AddBeforeSelf(created);
            }
            return;
        }
        existing.Element(ChartNs + "showVal")!.SetAttributeValue("val", labels.ShowValue ? "1" : "0");
        existing.Element(ChartNs + "showCatName")!.SetAttributeValue("val", labels.ShowCategoryName ? "1" : "0");
    }

    internal static string Semantics(SpreadsheetChartDataLabelsArtifact? labels) => labels is null
        ? "-"
        : $"value:{(labels.ShowValue ? 1 : 0)};category:{(labels.ShowCategoryName ? 1 : 0)}";

    private static XElement BooleanElement(string name, bool value) =>
        new(ChartNs + name, new XAttribute("val", value ? "1" : "0"));

    private static bool TryBoolean(XElement? element, out bool value)
    {
        value = false;
        if (element is null || element.Elements().Any() || HasUnexpectedText(element)) return false;
        var attributes = element.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToArray();
        if (attributes.Length != 1 || attributes[0].Name != "val" || !BooleanValues.Contains(attributes[0].Value)) return false;
        value = attributes[0].Value is "1" or "true";
        return true;
    }

    private static bool HasUnexpectedText(XElement element) => element.Nodes().Any(node => node switch
    {
        XElement => false,
        XText text => !string.IsNullOrWhiteSpace(text.Value),
        _ => true,
    });
}
