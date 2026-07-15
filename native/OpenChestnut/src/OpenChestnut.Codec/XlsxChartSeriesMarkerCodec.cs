using System.Globalization;
using System.Xml.Linq;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns the direct symbol/size plus bounded RGB fill/outline projection of one
// line-series marker. Picture markers, transformed colors, extensions, and
// unknown children keep the containing source-bound chart read-only so their
// native graph stays exact.
internal static class XlsxChartSeriesMarkerCodec
{
    private const uint MinSize = 2;
    private const uint MaxSize = 72;
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    internal static void Validate(SpreadsheetChartSeriesArtifact series, SpreadsheetChartType chartType, string worksheetId, string chartId)
    {
        var marker = series.Marker;
        if (marker is null) return;
        if (chartType != SpreadsheetChartType.Line)
            throw Invalid(worksheetId, chartId, series.Name, "is supported only on line charts");
        if (!IsSupported(marker.Symbol))
            throw Invalid(worksheetId, chartId, series.Name, "symbol is outside the bounded catalog");
        if (marker.HasSize && marker.Size is < MinSize or > MaxSize)
            throw Invalid(worksheetId, chartId, series.Name, $"size must be an integer from {MinSize} through {MaxSize}");
        XlsxChartSeriesStyleCodec.ValidateFill(marker.Fill, worksheetId, chartId, series.Name, "marker fill");
        XlsxChartSeriesLineStyleCodec.ValidateLine(marker.Line, worksheetId, chartId, series.Name, "marker line");
    }

    internal static bool TryRead(XElement nativeSeries, SpreadsheetChartSeriesArtifact series, SpreadsheetChartType chartType)
    {
        var markers = nativeSeries.Elements(ChartNs + "marker").ToArray();
        if (markers.Length == 0) return true;
        if (markers.Length != 1 || chartType != SpreadsheetChartType.Line) return false;
        var native = markers[0];
        if (native.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) ||
            native.Nodes().Any(node => node is XText text ? !string.IsNullOrWhiteSpace(text.Value) : node is not XElement)) return false;
        var children = native.Elements().ToArray();
        if (children.Any(child => child.Name != ChartNs + "symbol" && child.Name != ChartNs + "size" && child.Name != ChartNs + "spPr") ||
            children.Count(child => child.Name == ChartNs + "symbol") > 1 ||
            children.Count(child => child.Name == ChartNs + "size") > 1 ||
            children.Count(child => child.Name == ChartNs + "spPr") > 1) return false;

        var output = new SpreadsheetChartMarkerArtifact();
        var symbol = native.Element(ChartNs + "symbol");
        if (symbol is not null)
        {
            if (!IsScalar(symbol) || !TrySymbol((string?)symbol.Attribute("val"), out var value)) return false;
            output.Symbol = value;
        }
        var size = native.Element(ChartNs + "size");
        if (size is not null)
        {
            if (!IsScalar(size) || !uint.TryParse((string?)size.Attribute("val"), NumberStyles.None, CultureInfo.InvariantCulture, out var value) || value is < MinSize or > MaxSize) return false;
            output.Size = value;
        }
        var shapeProperties = native.Element(ChartNs + "spPr");
        if (shapeProperties is not null)
        {
            if (shapeProperties.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) ||
                shapeProperties.Nodes().Any(node => node is XText text ? !string.IsNullOrWhiteSpace(text.Value) : node is not XElement)) return false;
            var shapeChildren = shapeProperties.Elements().ToArray();
            if (shapeChildren.Any(child => child.Name != DrawingNs + "solidFill" && child.Name != DrawingNs + "ln") ||
                shapeChildren.Count(child => child.Name == DrawingNs + "solidFill") > 1 ||
                shapeChildren.Count(child => child.Name == DrawingNs + "ln") > 1) return false;
            if (!XlsxChartSeriesStyleCodec.TryReadSolidFill(shapeProperties, out var fill) ||
                !XlsxChartSeriesLineStyleCodec.TryReadLine(shapeProperties, out var line)) return false;
            if (fill is not null) output.Fill = fill;
            if (line is not null) output.Line = line;
        }
        series.Marker = output;
        return true;
    }

    internal static XElement? Element(SpreadsheetChartMarkerArtifact? marker)
    {
        if (marker is null) return null;
        var output = new XElement(ChartNs + "marker");
        if (marker.Symbol != SpreadsheetChartMarkerSymbol.Unspecified)
            output.Add(new XElement(ChartNs + "symbol", new XAttribute("val", SymbolValue(marker.Symbol))));
        if (marker.HasSize)
            output.Add(new XElement(ChartNs + "size", new XAttribute("val", marker.Size.ToString(CultureInfo.InvariantCulture))));
        var fill = marker.Fill is null ? null : XlsxChartSeriesStyleCodec.SolidFillElement(marker.Fill.Rgb);
        var line = XlsxChartSeriesLineStyleCodec.Element(marker.Line);
        if (fill is not null || line is not null) output.Add(new XElement(ChartNs + "spPr", fill, line));
        return output;
    }

    internal static void Patch(XElement nativeSeries, SpreadsheetChartSeriesArtifact target)
    {
        var existing = nativeSeries.Element(ChartNs + "marker");
        if (target.Marker is null) { existing?.Remove(); return; }
        var replacement = Element(target.Marker)!;
        if (existing is not null) { existing.ReplaceWith(replacement); return; }
        var after = nativeSeries.Element(ChartNs + "spPr") ?? nativeSeries.Element(ChartNs + "tx") ?? nativeSeries.Element(ChartNs + "order") ?? nativeSeries.Element(ChartNs + "idx");
        if (after is null) nativeSeries.AddFirst(replacement);
        else after.AddAfterSelf(replacement);
    }

    internal static string Semantics(SpreadsheetChartMarkerArtifact? marker)
    {
        if (marker is null) return "no-marker";
        var fill = marker.Fill is null
            ? "no-fill"
            : string.Join(':', marker.Fill.SourceCase, marker.Fill.Rgb.ToUpperInvariant(), marker.Fill.HasTint ? marker.Fill.Tint.ToString("R", CultureInfo.InvariantCulture) : "no-tint");
        return string.Join(':', "marker", (int)marker.Symbol, marker.HasSize ? marker.Size.ToString(CultureInfo.InvariantCulture) : "no-size", fill, XlsxChartSeriesLineStyleCodec.Semantics(marker.Line));
    }

    private static bool IsScalar(XElement element) =>
        element.Attributes().All(attribute => attribute.IsNamespaceDeclaration || attribute.Name == "val") &&
        element.Nodes().All(node => node is XText text && string.IsNullOrWhiteSpace(text.Value));

    private static bool IsSupported(SpreadsheetChartMarkerSymbol value) => value is
        SpreadsheetChartMarkerSymbol.Unspecified or SpreadsheetChartMarkerSymbol.None or SpreadsheetChartMarkerSymbol.Dot or
        SpreadsheetChartMarkerSymbol.Circle or SpreadsheetChartMarkerSymbol.Square or SpreadsheetChartMarkerSymbol.Diamond or
        SpreadsheetChartMarkerSymbol.Triangle or SpreadsheetChartMarkerSymbol.X or SpreadsheetChartMarkerSymbol.Star or
        SpreadsheetChartMarkerSymbol.Plus or SpreadsheetChartMarkerSymbol.Dash;

    private static bool TrySymbol(string? value, out SpreadsheetChartMarkerSymbol symbol)
    {
        symbol = value switch
        {
            "none" => SpreadsheetChartMarkerSymbol.None,
            "dot" => SpreadsheetChartMarkerSymbol.Dot,
            "circle" => SpreadsheetChartMarkerSymbol.Circle,
            "square" => SpreadsheetChartMarkerSymbol.Square,
            "diamond" => SpreadsheetChartMarkerSymbol.Diamond,
            "triangle" => SpreadsheetChartMarkerSymbol.Triangle,
            "x" => SpreadsheetChartMarkerSymbol.X,
            "star" => SpreadsheetChartMarkerSymbol.Star,
            "plus" => SpreadsheetChartMarkerSymbol.Plus,
            "dash" => SpreadsheetChartMarkerSymbol.Dash,
            _ => SpreadsheetChartMarkerSymbol.Unspecified,
        };
        return symbol != SpreadsheetChartMarkerSymbol.Unspecified;
    }

    private static string SymbolValue(SpreadsheetChartMarkerSymbol symbol) => symbol switch
    {
        SpreadsheetChartMarkerSymbol.None => "none",
        SpreadsheetChartMarkerSymbol.Dot => "dot",
        SpreadsheetChartMarkerSymbol.Circle => "circle",
        SpreadsheetChartMarkerSymbol.Square => "square",
        SpreadsheetChartMarkerSymbol.Diamond => "diamond",
        SpreadsheetChartMarkerSymbol.Triangle => "triangle",
        SpreadsheetChartMarkerSymbol.X => "x",
        SpreadsheetChartMarkerSymbol.Star => "star",
        SpreadsheetChartMarkerSymbol.Plus => "plus",
        SpreadsheetChartMarkerSymbol.Dash => "dash",
        _ => throw new InvalidOperationException("Validated worksheet chart marker symbol changed unexpectedly."),
    };

    private static CodecException Invalid(string worksheetId, string chartId, string seriesName, string message) =>
        new("invalid_spreadsheet_chart", $"Worksheet {worksheetId} chart {chartId} series {seriesName} marker {message}.");
}
