using System.Globalization;
using System.Xml.Linq;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns the bounded series-level DrawingML style projection. The public model
// currently exposes only an explicit six-digit RGB solid fill. All other
// series shape properties remain source-owned residual XML; unrecognized fill
// kinds make the containing chart read-only instead of being flattened.
internal static class XlsxChartSeriesStyleCodec
{
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly HashSet<XName> FillNames =
    [
        DrawingNs + "noFill",
        DrawingNs + "solidFill",
        DrawingNs + "gradFill",
        DrawingNs + "blipFill",
        DrawingNs + "pattFill",
        DrawingNs + "grpFill",
    ];

    internal static void Validate(SpreadsheetChartSeriesArtifact series, string worksheetId, string chartId)
    {
        if (series.Fill is null) return;
        if (series.Fill.SourceCase != SpreadsheetColor.SourceOneofCase.Rgb ||
            series.Fill.HasTint ||
            series.Fill.Rgb.Length != 6 ||
            !series.Fill.Rgb.All(Uri.IsHexDigit))
        {
            throw new CodecException(
                "invalid_spreadsheet_chart",
                $"Worksheet {worksheetId} chart {chartId} series {series.Name} fill must be an untinted six-digit RGB solid color.");
        }
    }

    internal static bool TryRead(XElement nativeSeries, SpreadsheetChartSeriesArtifact series)
    {
        var shapeProperties = nativeSeries.Element(ChartNs + "spPr");
        if (shapeProperties is null) return true;
        var fills = shapeProperties.Elements().Where(item => FillNames.Contains(item.Name)).ToArray();
        if (fills.Length == 0) return true;
        if (fills.Length != 1 || fills[0].Name != DrawingNs + "solidFill") return false;

        var solidFill = fills[0];
        if (solidFill.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration)) return false;
        var colors = solidFill.Elements().ToArray();
        if (colors.Length != 1 || colors[0].Name != DrawingNs + "srgbClr") return false;
        var color = colors[0];
        if (color.HasElements || color.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != "val")) return false;
        var value = (string?)color.Attribute("val");
        if (value is null || value.Length != 6 || !value.All(Uri.IsHexDigit)) return false;
        series.Fill = new SpreadsheetColor { Rgb = value.ToUpperInvariant() };
        return true;
    }

    internal static XElement? PropertiesElement(SpreadsheetChartSeriesArtifact series)
    {
        var fill = series.Fill is null ? null : SolidFillElement(series.Fill.Rgb);
        var line = XlsxChartSeriesLineStyleCodec.Element(series.Line);
        return fill is null && line is null ? null : new XElement(ChartNs + "spPr", fill, line);
    }

    internal static void Patch(XElement nativeSeries, SpreadsheetChartSeriesArtifact target)
    {
        var shapeProperties = nativeSeries.Element(ChartNs + "spPr");
        var solidFill = shapeProperties?.Element(DrawingNs + "solidFill");
        if (target.Fill is null)
        {
            solidFill?.Remove();
            if (shapeProperties is not null && !shapeProperties.Elements().Any() && !shapeProperties.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration)) shapeProperties.Remove();
            return;
        }

        if (shapeProperties is null)
        {
            shapeProperties = new XElement(ChartNs + "spPr");
            var before = nativeSeries.Elements().FirstOrDefault(item => item.Name != ChartNs + "idx" && item.Name != ChartNs + "order" && item.Name != ChartNs + "tx");
            if (before is null) nativeSeries.Add(shapeProperties);
            else before.AddBeforeSelf(shapeProperties);
        }
        var replacement = SolidFillElement(target.Fill.Rgb);
        if (solidFill is not null) solidFill.ReplaceWith(replacement);
        else
        {
            var before = shapeProperties.Elements().FirstOrDefault(item => IsShapePropertyTail(item.Name));
            if (before is null) shapeProperties.Add(replacement);
            else before.AddBeforeSelf(replacement);
        }
    }

    internal static string Semantics(SpreadsheetChartSeriesArtifact series) =>
        series.Fill is null ? "no-fill" : string.Join(':', "rgb", series.Fill.Rgb.ToUpperInvariant(), series.Fill.HasTint ? series.Fill.Tint.ToString("R", CultureInfo.InvariantCulture) : "no-tint");

    private static XElement SolidFillElement(string rgb) =>
        new(DrawingNs + "solidFill", new XElement(DrawingNs + "srgbClr", new XAttribute("val", rgb.ToUpperInvariant())));

    private static bool IsShapePropertyTail(XName name) =>
        name == DrawingNs + "ln" || name == DrawingNs + "effectLst" || name == DrawingNs + "effectDag" ||
        name == DrawingNs + "scene3d" || name == DrawingNs + "sp3d" || name == DrawingNs + "extLst";
}
