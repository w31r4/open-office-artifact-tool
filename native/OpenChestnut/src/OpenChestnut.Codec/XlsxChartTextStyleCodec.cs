using System.Globalization;
using System.Xml.Linq;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns only explicit point sizes for the chart title and primary-axis tick
// labels. Rich DrawingML text graphs are deliberately not normalized: an
// unrecognized title run or txPr makes the containing chart read-only while the
// hash-bound ChartPart remains byte-exact on an unchanged export.
internal static class XlsxChartTextStyleCodec
{
    private const double MinimumFontSizePoints = 1;
    private const double MaximumFontSizePoints = 4_000;
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    internal static void Validate(SpreadsheetChartArtifact chart, string worksheetId)
    {
        ValidateStyle(chart.TitleTextStyle, worksheetId, chart.Id, "title_text_style");
        if (chart.TitleTextStyle is not null && chart.Title.Length == 0)
            throw Invalid(worksheetId, chart.Id, "title_text_style requires a non-empty title.");
        ValidateStyle(chart.XAxis?.TextStyle, worksheetId, chart.Id, "x_axis.text_style");
        ValidateStyle(chart.YAxis?.TextStyle, worksheetId, chart.Id, "y_axis.text_style");
    }

    internal static bool TryReadTitle(XElement title, SpreadsheetChartArtifact chart)
    {
        if (!TryExactTitleRun(title, out var run)) return false;
        var runProperties = run.Element(DrawingNs + "rPr");
        if (runProperties is null) return true;
        if (!TryExactSizeProperties(runProperties, out var fontSize)) return false;
        chart.TitleTextStyle = new SpreadsheetChartTextStyleArtifact { FontSizePoints = fontSize };
        return true;
    }

    internal static bool TryReadAxis(XElement axis, SpreadsheetChartAxisArtifact semantic)
    {
        var properties = axis.Elements(ChartNs + "txPr").Take(2).ToArray();
        if (properties.Length == 0) return true;
        if (properties.Length != 1 || !TryExactAxisTextProperties(properties[0], out var fontSize)) return false;
        semantic.TextStyle = new SpreadsheetChartTextStyleArtifact { FontSizePoints = fontSize };
        return true;
    }

    internal static XElement TitleElement(string title, SpreadsheetChartTextStyleArtifact? style)
    {
        var run = new XElement(DrawingNs + "r");
        if (style is not null) run.Add(SizeProperties("rPr", style.FontSizePoints));
        run.Add(new XElement(DrawingNs + "t", title));
        return new XElement(ChartNs + "title",
            new XElement(ChartNs + "tx", new XElement(ChartNs + "rich",
                new XElement(DrawingNs + "bodyPr"), new XElement(DrawingNs + "lstStyle"),
                new XElement(DrawingNs + "p", run))),
            new XElement(ChartNs + "layout"));
    }

    internal static void AppendAuthoredAxis(XElement axis, SpreadsheetChartTextStyleArtifact? style)
    {
        if (style is not null) axis.Add(AxisTextProperties(style.FontSizePoints));
    }

    internal static void PatchTitle(XElement title, SpreadsheetChartTextStyleArtifact? style)
    {
        if (!TryExactTitleRun(title, out var run))
            throw new CodecException("unsupported_spreadsheet_chart_edit", "Referenced worksheet-chart title text styling is read-only.");
        PatchSizeProperties(run, "rPr", style);
    }

    internal static void PatchAxis(XElement axis, SpreadsheetChartTextStyleArtifact? style)
    {
        var existing = axis.Element(ChartNs + "txPr");
        if (style is null) { existing?.Remove(); return; }
        if (existing is null)
        {
            var created = AxisTextProperties(style.FontSizePoints);
            var crossAxis = axis.Element(ChartNs + "crossAx");
            if (crossAxis is null) axis.Add(created);
            else crossAxis.AddBeforeSelf(created);
            return;
        }
        if (!TryExactAxisTextProperties(existing, out _))
            throw new CodecException("unsupported_spreadsheet_chart_edit", "Referenced worksheet-chart axis text styling is read-only.");
        existing.Descendants(DrawingNs + "defRPr").Single().SetAttributeValue("sz", Size(style.FontSizePoints));
    }

    internal static string Semantics(SpreadsheetChartTextStyleArtifact? style) =>
        style is null ? "-" : style.HasFontSizePoints ? style.FontSizePoints.ToString("R", CultureInfo.InvariantCulture) : "present-without-size";

    private static void ValidateStyle(SpreadsheetChartTextStyleArtifact? style, string worksheetId, string chartId, string field)
    {
        if (style is null) return;
        if (!style.HasFontSizePoints || !double.IsFinite(style.FontSizePoints) || style.FontSizePoints < MinimumFontSizePoints || style.FontSizePoints > MaximumFontSizePoints)
            throw Invalid(worksheetId, chartId, $"{field}.font_size_points must be from 1 through 4000.");
    }

    private static bool TryExactTitleRun(XElement title, out XElement run)
    {
        run = null!;
        var titleChildren = title.Elements().ToArray();
        if (titleChildren.Any(item => item.Name != ChartNs + "tx" && item.Name != ChartNs + "layout") || titleChildren.Count(item => item.Name == ChartNs + "tx") != 1) return false;
        var tx = title.Element(ChartNs + "tx")!;
        var rich = tx.Element(ChartNs + "rich");
        if (rich is null || tx.Elements().Count() != 1) return false;
        var richChildren = rich.Elements().ToArray();
        if (richChildren.Length != 3 || richChildren[0].Name != DrawingNs + "bodyPr" || richChildren[1].Name != DrawingNs + "lstStyle" || richChildren[2].Name != DrawingNs + "p" ||
            richChildren[0].HasAttributes || richChildren[0].HasElements || richChildren[1].HasAttributes || richChildren[1].HasElements) return false;
        var paragraph = richChildren[2];
        var paragraphChildren = paragraph.Elements().ToArray();
        if (paragraphChildren.Length != 1 || paragraphChildren[0].Name != DrawingNs + "r") return false;
        run = paragraphChildren[0];
        var runChildren = run.Elements().ToArray();
        if (runChildren.Length is < 1 or > 2 || runChildren[^1].Name != DrawingNs + "t" ||
            runChildren.Length == 2 && runChildren[0].Name != DrawingNs + "rPr") return false;
        return true;
    }

    private static bool TryExactAxisTextProperties(XElement properties, out double fontSize)
    {
        fontSize = 0;
        var children = properties.Elements().ToArray();
        if (children.Length != 3 || children[0].Name != DrawingNs + "bodyPr" || children[1].Name != DrawingNs + "lstStyle" || children[2].Name != DrawingNs + "p" ||
            children[0].HasAttributes || children[0].HasElements || children[1].HasAttributes || children[1].HasElements) return false;
        var paragraphChildren = children[2].Elements().ToArray();
        if (paragraphChildren.Length != 2 || paragraphChildren[0].Name != DrawingNs + "pPr" || paragraphChildren[1].Name != DrawingNs + "endParaRPr" ||
            paragraphChildren[1].HasAttributes || paragraphChildren[1].HasElements) return false;
        var paragraphProperties = paragraphChildren[0];
        var defaults = paragraphProperties.Elements().ToArray();
        return !paragraphProperties.HasAttributes && defaults.Length == 1 && defaults[0].Name == DrawingNs + "defRPr" && TryExactSizeProperties(defaults[0], out fontSize);
    }

    private static bool TryExactSizeProperties(XElement properties, out double fontSize)
    {
        fontSize = 0;
        var attributes = properties.Attributes().ToArray();
        if (attributes.Length != 1 || attributes[0].Name != "sz" || properties.HasElements ||
            !uint.TryParse(attributes[0].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var size) || size is < 100 or > 400_000) return false;
        fontSize = size / 100d;
        return true;
    }

    private static void PatchSizeProperties(XElement owner, string name, SpreadsheetChartTextStyleArtifact? style)
    {
        var existing = owner.Element(DrawingNs + name);
        if (style is null) { existing?.Remove(); return; }
        if (existing is null) owner.AddFirst(SizeProperties(name, style.FontSizePoints));
        else
        {
            if (!TryExactSizeProperties(existing, out _)) throw new CodecException("unsupported_spreadsheet_chart_edit", "Referenced worksheet-chart title text styling is read-only.");
            existing.SetAttributeValue("sz", Size(style.FontSizePoints));
        }
    }

    private static XElement AxisTextProperties(double fontSize) => new(ChartNs + "txPr",
        new XElement(DrawingNs + "bodyPr"), new XElement(DrawingNs + "lstStyle"),
        new XElement(DrawingNs + "p",
            new XElement(DrawingNs + "pPr", SizeProperties("defRPr", fontSize)),
            new XElement(DrawingNs + "endParaRPr")));

    private static XElement SizeProperties(string name, double fontSize) => new(DrawingNs + name, new XAttribute("sz", Size(fontSize)));
    private static uint Size(double points) => checked((uint)Math.Round(points * 100, MidpointRounding.AwayFromZero));
    private static CodecException Invalid(string worksheetId, string chartId, string message) => new("invalid_spreadsheet_chart", $"Worksheet {worksheetId} chart {chartId} {message}");
}
