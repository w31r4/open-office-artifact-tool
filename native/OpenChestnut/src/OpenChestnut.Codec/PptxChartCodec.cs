using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenChestnut.Codec;

// Owns the bounded literal-data p:graphicFrame -> ChartPart projection. The
// chart semantic atoms deliberately reuse the worksheet-chart wire messages;
// PresentationML contributes only its page-relative frame.
internal static partial class PptxChartCodec
{
    private const string ChartGraphicDataUri = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private const int MaxSeries = 256;
    private const int MaxPoints = 1_048_576;
    private static readonly XNamespace ChartNs = ChartGraphicDataUri;
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    internal sealed record Replacement(string PartPath, string Sha256);

    internal static bool TryRead(P.GraphicFrame source, PptxPartContext context, out PresentationChart chart, out bool editable)
    {
        chart = new PresentationChart();
        editable = false;
        try
        {
            if (source.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties is not { Id.HasValue: true } ||
                source.Transform is not { } transform || !TryReadFrame(transform, out var left, out var top, out var width, out var height) ||
                source.Graphic?.GraphicData is not { } graphicData ||
                !string.Equals(graphicData.Uri?.Value, ChartGraphicDataUri, StringComparison.Ordinal) ||
                graphicData.Elements<C.ChartReference>().SingleOrDefault()?.Id?.Value is not { Length: > 0 } relationshipId)
                return false;
            ChartPart part;
            try
            {
                if (context.Owner.GetPartById(relationshipId) is not ChartPart chartPart) return false;
                part = chartPart;
            }
            catch (ArgumentOutOfRangeException)
            {
                return false;
            }
            var xml = ReadXml(part);
            if (TryReadComboChart(xml, out chart, out _, out editable))
            {
                chart.LeftEmu = left;
                chart.TopEmu = top;
                chart.WidthEmu = width;
                chart.HeightEmu = height;
                return true;
            }
            if (!TryReadChart(xml, out var semantic, out _, out editable)) return false;
            chart = FromSpreadsheet(semantic, left, top, width, height);
            return true;
        }
        catch (Exception error) when (error is InvalidOperationException or OverflowException or XmlException)
        {
            chart = new PresentationChart();
            editable = false;
            return false;
        }
    }

    internal static P.GraphicFrame Build(PresentationElement element, uint nativeId, SlidePart slidePart)
    {
        Validate(element.Chart, element.Id, element.Name);
        var chartPart = slidePart.AddNewPart<ChartPart>();
        WriteXml(chartPart, BuildPresentationChartDocument(element.Chart, element.Id, element.Name));
        var relationshipId = slidePart.GetIdOfPart(chartPart);
        return new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = nativeId, Name = element.Name },
                new P.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.Transform(
                new A.Offset { X = element.Chart.LeftEmu, Y = element.Chart.TopEmu },
                new A.Extents { Cx = element.Chart.WidthEmu, Cy = element.Chart.HeightEmu }),
            new A.Graphic(new A.GraphicData(new C.ChartReference { Id = relationshipId }) { Uri = ChartGraphicDataUri }));
    }

    internal static Replacement Apply(P.GraphicFrame source, PresentationElement requested, PptxPartContext context)
    {
        if (!TryRead(source, context, out var original, out var editable) || !editable)
            throw new CodecException("unsupported_presentation_edit", $"Presentation chart {requested.Id} no longer matches the editable literal-data chart profile.");
        Validate(requested.Chart, requested.Id, requested.Name);
        if (!PresentationChartTopologyMatches(requested.Chart, original))
            throw new CodecException("presentation_chart_topology_changed", $"Presentation chart {requested.Id} cannot change chart type, series count, or point topology.");

        var relationshipId = source.Graphic!.GraphicData!.Elements<C.ChartReference>().Single().Id!.Value!;
        var part = (ChartPart)context.Owner.GetPartById(relationshipId);
        var document = XDocument.Parse(ReadXml(part), LoadOptions.PreserveWhitespace);
        PatchPresentationChart(document, requested.Chart, requested.Id, requested.Name);
        WriteXml(part, document);
        source.NonVisualGraphicFrameProperties!.NonVisualDrawingProperties!.Name = requested.Name;
        SetFrame(source.Transform!, requested.Chart);
        var bytes = ReadBytes(part);
        return new Replacement(Path(part), Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant());
    }

    internal static void Validate(PresentationChart? chart, string elementId, string name)
    {
        if (chart is null) throw Invalid(elementId, "payload is missing");
        if (string.IsNullOrWhiteSpace(name) || name.Length > 255 || HasControls(name)) throw Invalid(elementId, "name must contain 1 through 255 characters without controls");
        if (chart.LeftEmu < 0 || chart.TopEmu < 0 || chart.WidthEmu <= 0 || chart.HeightEmu <= 0)
            throw Invalid(elementId, "frame must have non-negative coordinates and positive dimensions");
        if (chart.Type == SpreadsheetChartType.Combo)
        {
            ValidateComboChart(chart, elementId, name);
            return;
        }
        if (chart.ComboSeries.Count != 0) throw Invalid(elementId, "must not carry combo_series unless type is combo");
        if (chart.Series.Any(series => !string.IsNullOrWhiteSpace(series.CategoryFormula) || !string.IsNullOrWhiteSpace(series.XValueFormula) || !string.IsNullOrWhiteSpace(series.ValueFormula) || !string.IsNullOrWhiteSpace(series.BubbleSizeFormula)))
            throw Invalid(elementId, "must use literal categories and values without workbook formulas");
        var spreadsheet = ToSpreadsheet(chart, elementId, name);
        try
        {
            XlsxChartCodec.Validate([spreadsheet], $"presentation/{elementId}");
        }
        catch (CodecException error) when (error.Code == "invalid_spreadsheet_chart")
        {
            throw Invalid(elementId, error.Message);
        }
    }

    internal static void ScrubFrame(P.GraphicFrame source)
    {
        if (source.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties is { } nonVisual) nonVisual.Name = string.Empty;
        if (source.Transform is { } transform)
        {
            transform.Offset!.X = 0L;
            transform.Offset.Y = 0L;
            transform.Extents!.Cx = 1L;
            transform.Extents.Cy = 1L;
        }
    }

    private static SpreadsheetChartArtifact ToSpreadsheet(PresentationChart source, string id, string name)
    {
        var output = new SpreadsheetChartArtifact
        {
            Id = id,
            Name = name,
            Title = source.Title,
            Type = source.Type,
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
        output.Series.Add(source.Series.Select(series => series.Clone()));
        if (source.XAxis is not null) output.XAxis = source.XAxis.Clone();
        if (source.YAxis is not null) output.YAxis = source.YAxis.Clone();
        if (source.DataLabels is not null) output.DataLabels = source.DataLabels.Clone();
        return output;
    }

    private static PresentationChart FromSpreadsheet(SpreadsheetChartArtifact source, long left, long top, long width, long height)
    {
        var output = new PresentationChart
        {
            LeftEmu = left,
            TopEmu = top,
            WidthEmu = width,
            HeightEmu = height,
            Type = source.Type,
            Title = source.Title,
            HasLegend = source.HasLegend,
        };
        output.Categories.Add(source.Categories);
        output.Series.Add(source.Series.Select(series => series.Clone()));
        if (source.XAxis is not null) output.XAxis = source.XAxis.Clone();
        if (source.YAxis is not null) output.YAxis = source.YAxis.Clone();
        if (source.DataLabels is not null) output.DataLabels = source.DataLabels.Clone();
        return output;
    }

    private static bool TryReadChart(string xml, out SpreadsheetChartArtifact chart, out XDocument document, out bool editable)
    {
        if (!OpenXmlChartSpaceCodec.TryRead(xml, out chart, out document, out editable)) return false;
        return chart.Series.All(series =>
            string.IsNullOrWhiteSpace(series.CategoryFormula) &&
            string.IsNullOrWhiteSpace(series.XValueFormula) &&
            string.IsNullOrWhiteSpace(series.ValueFormula) &&
            string.IsNullOrWhiteSpace(series.BubbleSizeFormula));
    }

    private static XDocument BuildChartDocument(SpreadsheetChartArtifact chart)
    {
        return OpenXmlChartSpaceCodec.Build(chart);
    }

    private static void PatchChart(XDocument document, SpreadsheetChartArtifact target)
    {
        OpenXmlChartSpaceCodec.Patch(document, target, "presentation_chart_topology_changed", "Presentation chart");
    }

    private static bool TryReadFrame(P.Transform transform, out long left, out long top, out long width, out long height)
    {
        left = top = width = height = 0;
        if (transform.Offset?.X?.Value is null || transform.Offset.Y?.Value is null || transform.Extents?.Cx?.Value is null or <= 0 || transform.Extents.Cy?.Value is null or <= 0 || transform.Offset.X.Value < 0 || transform.Offset.Y.Value < 0) return false;
        left = transform.Offset.X.Value; top = transform.Offset.Y.Value; width = transform.Extents.Cx.Value; height = transform.Extents.Cy.Value;
        return true;
    }

    private static void SetFrame(P.Transform transform, PresentationChart chart)
    {
        transform.Offset!.X = chart.LeftEmu; transform.Offset.Y = chart.TopEmu;
        transform.Extents!.Cx = chart.WidthEmu; transform.Extents.Cy = chart.HeightEmu;
    }

    private static string ReadXml(OpenXmlPart part) => Encoding.UTF8.GetString(ReadBytes(part));
    private static byte[] ReadBytes(OpenXmlPart part) { using var stream = part.GetStream(FileMode.Open, FileAccess.Read); using var memory = new MemoryStream(); stream.CopyTo(memory); return memory.ToArray(); }
    private static void WriteXml(OpenXmlPart part, XDocument document) { using var stream = part.GetStream(FileMode.Create, FileAccess.Write); using var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false), OmitXmlDeclaration = false, Indent = false }); document.Save(writer); }
    private static string Path(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static bool HasControls(string value) => value.Any(char.IsControl);
    private static CodecException Invalid(string id, string message) => new("invalid_presentation_chart", $"Presentation chart {id} {message}.");
}
