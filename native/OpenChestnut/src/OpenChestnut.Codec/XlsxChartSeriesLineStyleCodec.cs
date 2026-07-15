using System.Globalization;
using System.Xml.Linq;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns one bounded direct DrawingML series outline. The projection is presence
// aware and deliberately excludes theme/transformed colors, compound lines,
// caps, joins, arrows, custom dash arrays, and other line children. Encountering
// any of those graphs makes the containing chart read-only and exact-preserved.
internal static class XlsxChartSeriesLineStyleCodec
{
    private const double MaxWidthPoints = 1_584;
    private const long EmuPerPoint = 12_700;
    private const long MaxWidthEmu = 20_116_800;
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    internal static void Validate(SpreadsheetChartSeriesArtifact series, string worksheetId, string chartId)
    {
        var line = series.Line;
        if (line is null) return;
        if (line.Color is not null &&
            (line.Color.SourceCase != SpreadsheetColor.SourceOneofCase.Rgb || line.Color.HasTint ||
             line.Color.Rgb.Length != 6 || !line.Color.Rgb.All(Uri.IsHexDigit)))
            throw Invalid(worksheetId, chartId, series.Name, "color must be an untinted six-digit RGB value");
        if (line.DashStyle is not (SpreadsheetChartLineDashStyle.Unspecified or SpreadsheetChartLineDashStyle.Solid or
            SpreadsheetChartLineDashStyle.Dashed or SpreadsheetChartLineDashStyle.Dotted or
            SpreadsheetChartLineDashStyle.DashDot or SpreadsheetChartLineDashStyle.DashDotDot))
            throw Invalid(worksheetId, chartId, series.Name, "dash style is outside the bounded preset catalog");
        if (line.HasWidthPoints &&
            (double.IsNaN(line.WidthPoints) || double.IsInfinity(line.WidthPoints) ||
             line.WidthPoints < 0 || line.WidthPoints > MaxWidthPoints || WidthEmu(line.WidthPoints) > MaxWidthEmu))
            throw Invalid(worksheetId, chartId, series.Name, $"width must be from 0 through {MaxWidthPoints} points");
    }

    internal static bool TryRead(XElement nativeSeries, SpreadsheetChartSeriesArtifact series)
    {
        var shapeProperties = nativeSeries.Element(ChartNs + "spPr");
        if (shapeProperties is null) return true;
        var lines = shapeProperties.Elements(DrawingNs + "ln").ToArray();
        if (lines.Length == 0) return true;
        if (lines.Length != 1) return false;
        var native = lines[0];
        if (native.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != "w")) return false;
        var children = native.Elements().ToArray();
        if (children.Any(child => child.Name != DrawingNs + "solidFill" && child.Name != DrawingNs + "prstDash") ||
            children.Count(child => child.Name == DrawingNs + "solidFill") > 1 ||
            children.Count(child => child.Name == DrawingNs + "prstDash") > 1) return false;

        var output = new SpreadsheetChartLineStyleArtifact();
        var width = (string?)native.Attribute("w");
        if (width is not null)
        {
            if (!long.TryParse(width, NumberStyles.None, CultureInfo.InvariantCulture, out var emu) || emu < 0 || emu > MaxWidthEmu) return false;
            output.WidthPoints = emu / (double)EmuPerPoint;
        }
        var fill = native.Element(DrawingNs + "solidFill");
        if (fill is not null)
        {
            if (fill.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration)) return false;
            var colors = fill.Elements().ToArray();
            if (colors.Length != 1 || colors[0].Name != DrawingNs + "srgbClr") return false;
            var color = colors[0];
            if (color.HasElements || color.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != "val")) return false;
            var value = (string?)color.Attribute("val");
            if (value is null || value.Length != 6 || !value.All(Uri.IsHexDigit)) return false;
            output.Color = new SpreadsheetColor { Rgb = value.ToUpperInvariant() };
        }
        var dash = native.Element(DrawingNs + "prstDash");
        if (dash is not null)
        {
            if (dash.HasElements || dash.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != "val") || !TryDash((string?)dash.Attribute("val"), out var style)) return false;
            output.DashStyle = style;
        }
        series.Line = output;
        return true;
    }

    internal static XElement? Element(SpreadsheetChartLineStyleArtifact? line)
    {
        if (line is null) return null;
        var output = new XElement(DrawingNs + "ln");
        if (line.HasWidthPoints) output.SetAttributeValue("w", WidthEmu(line.WidthPoints).ToString(CultureInfo.InvariantCulture));
        if (line.Color is not null) output.Add(new XElement(DrawingNs + "solidFill", new XElement(DrawingNs + "srgbClr", new XAttribute("val", line.Color.Rgb.ToUpperInvariant()))));
        if (line.DashStyle != SpreadsheetChartLineDashStyle.Unspecified) output.Add(new XElement(DrawingNs + "prstDash", new XAttribute("val", DashValue(line.DashStyle))));
        return output;
    }

    internal static void Patch(XElement nativeSeries, SpreadsheetChartSeriesArtifact target)
    {
        var shapeProperties = nativeSeries.Element(ChartNs + "spPr");
        var existing = shapeProperties?.Element(DrawingNs + "ln");
        if (target.Line is null)
        {
            existing?.Remove();
            RemoveEmpty(shapeProperties);
            return;
        }
        if (shapeProperties is null)
        {
            shapeProperties = new XElement(ChartNs + "spPr");
            var before = nativeSeries.Elements().FirstOrDefault(item => item.Name != ChartNs + "idx" && item.Name != ChartNs + "order" && item.Name != ChartNs + "tx");
            if (before is null) nativeSeries.Add(shapeProperties);
            else before.AddBeforeSelf(shapeProperties);
        }
        var replacement = Element(target.Line)!;
        if (existing is not null) existing.ReplaceWith(replacement);
        else
        {
            var before = shapeProperties.Elements().FirstOrDefault(item => IsShapePropertyTail(item.Name));
            if (before is null) shapeProperties.Add(replacement);
            else before.AddBeforeSelf(replacement);
        }
    }

    internal static string Semantics(SpreadsheetChartLineStyleArtifact? line)
    {
        if (line is null) return "no-line";
        var color = line.Color is null ? "no-color" : string.Join(':', line.Color.SourceCase, line.Color.Rgb.ToUpperInvariant(), line.Color.HasTint ? line.Color.Tint.ToString("R", CultureInfo.InvariantCulture) : "no-tint");
        return string.Join(':', "line", color, (int)line.DashStyle, line.HasWidthPoints ? line.WidthPoints.ToString("R", CultureInfo.InvariantCulture) : "no-width");
    }

    private static long WidthEmu(double points) => checked((long)Math.Round(points * EmuPerPoint, MidpointRounding.AwayFromZero));

    private static bool TryDash(string? value, out SpreadsheetChartLineDashStyle style)
    {
        style = value switch
        {
            "solid" => SpreadsheetChartLineDashStyle.Solid,
            "dash" => SpreadsheetChartLineDashStyle.Dashed,
            "dot" => SpreadsheetChartLineDashStyle.Dotted,
            "dashDot" => SpreadsheetChartLineDashStyle.DashDot,
            "lgDashDotDot" => SpreadsheetChartLineDashStyle.DashDotDot,
            _ => SpreadsheetChartLineDashStyle.Unspecified,
        };
        return style != SpreadsheetChartLineDashStyle.Unspecified;
    }

    private static string DashValue(SpreadsheetChartLineDashStyle style) => style switch
    {
        SpreadsheetChartLineDashStyle.Solid => "solid",
        SpreadsheetChartLineDashStyle.Dashed => "dash",
        SpreadsheetChartLineDashStyle.Dotted => "dot",
        SpreadsheetChartLineDashStyle.DashDot => "dashDot",
        SpreadsheetChartLineDashStyle.DashDotDot => "lgDashDotDot",
        _ => throw new InvalidOperationException("Validated worksheet chart line dash style changed unexpectedly."),
    };

    private static bool IsShapePropertyTail(XName name) =>
        name == DrawingNs + "effectLst" || name == DrawingNs + "effectDag" || name == DrawingNs + "scene3d" ||
        name == DrawingNs + "sp3d" || name == DrawingNs + "extLst";

    private static void RemoveEmpty(XElement? shapeProperties)
    {
        if (shapeProperties is not null && !shapeProperties.Elements().Any() && !shapeProperties.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration)) shapeProperties.Remove();
    }

    private static CodecException Invalid(string worksheetId, string chartId, string seriesName, string message) =>
        new("invalid_spreadsheet_chart", $"Worksheet {worksheetId} chart {chartId} series {seriesName} line {message}.");
}
