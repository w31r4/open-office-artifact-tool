using System.Globalization;
using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Owns the first bounded direct table-formatting profile. The profile is
// deliberately complete: presence means fixed dxa width/grid, four cell
// margins, six uniform RGB borders, and a uniformly filled/bold first row.
// Partial, theme-based, conditional, or style-effective formatting remains in
// source XML instead of being represented as safely editable semantics.
internal static class DocxTableFormatting
{
    private const uint MaxDxa = 1_000_000;
    private const string WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    internal static void Validate(DocumentTable table)
    {
        var formatting = table.Formatting;
        if (formatting is null) return;
        var columns = LogicalColumns(table);
        if (formatting.WidthDxa is 0 or > MaxDxa)
            throw Invalid("Document table formatting width_dxa must be between 1 and 1000000.");
        if (formatting.IndentDxa > MaxDxa)
            throw Invalid("Document table formatting indent_dxa must not exceed 1000000.");
        if (formatting.ColumnWidthsDxa.Count != columns)
            throw Invalid($"Document table formatting requires one column_widths_dxa entry for each of its {columns} logical grid columns.");
        if (formatting.ColumnWidthsDxa.Any(width => width is 0 or > MaxDxa))
            throw Invalid("Document table formatting column widths must be between 1 and 1000000.");
        if (formatting.ColumnWidthsDxa.Aggregate(0UL, (sum, width) => checked(sum + width)) != formatting.WidthDxa)
            throw Invalid("Document table formatting column widths must sum exactly to width_dxa.");
        if (formatting.CellMarginsDxa is null)
            throw Invalid("Document table formatting requires all four cell_margins_dxa values.");
        if (new[]
            {
                formatting.CellMarginsDxa.Top,
                formatting.CellMarginsDxa.Bottom,
                formatting.CellMarginsDxa.Start,
                formatting.CellMarginsDxa.End,
            }.Any(margin => margin > MaxDxa))
            throw Invalid("Document table formatting cell margins must not exceed 1000000.");
        if (!IsRgb(formatting.BorderColor))
            throw Invalid("Document table formatting border_color must be a six-digit uppercase RGB value.");
        if (!IsRgb(formatting.HeaderFill))
            throw Invalid("Document table formatting header_fill must be a six-digit uppercase RGB value.");
        if (formatting.BorderSize is 1 or > 96)
            throw Invalid("Document table formatting border_size must be zero or between 2 and 96 eighths of a point.");
    }

    internal static W.TableProperties? BuildProperties(string styleId, DocumentTableFormatting? formatting)
    {
        if (string.IsNullOrWhiteSpace(styleId) && formatting is null) return null;
        var properties = new W.TableProperties();
        if (!string.IsNullOrWhiteSpace(styleId)) properties.Append(new W.TableStyle { Val = styleId });
        if (formatting is null) return properties;

        properties.Append(Dxa(new W.TableWidth(), formatting.WidthDxa));
        properties.Append(Dxa(new W.TableIndentation(), formatting.IndentDxa));
        properties.Append(new W.TableBorders(
            Border(new W.TopBorder(), formatting),
            Border(new W.LeftBorder(), formatting),
            Border(new W.BottomBorder(), formatting),
            Border(new W.RightBorder(), formatting),
            Border(new W.InsideHorizontalBorder(), formatting),
            Border(new W.InsideVerticalBorder(), formatting)));
        properties.Append(new W.TableLayout { Type = W.TableLayoutValues.Fixed });
        properties.Append(new W.TableCellMarginDefault(
            Dxa(new W.TopMargin(), formatting.CellMarginsDxa.Top),
            Dxa(new W.StartMargin(), formatting.CellMarginsDxa.Start),
            Dxa(new W.BottomMargin(), formatting.CellMarginsDxa.Bottom),
            Dxa(new W.EndMargin(), formatting.CellMarginsDxa.End)));
        return properties;
    }

    internal static W.GridColumn BuildGridColumn(DocumentTableFormatting? formatting, int column)
    {
        var gridColumn = new W.GridColumn();
        if (formatting is not null) SetWordAttribute(gridColumn, "w", formatting.ColumnWidthsDxa[column].ToString(CultureInfo.InvariantCulture));
        return gridColumn;
    }

    internal static W.TableCellWidth? BuildCellWidth(DocumentTableFormatting? formatting, uint gridColumn, uint columnSpan)
    {
        if (formatting is null) return null;
        var end = checked(gridColumn + columnSpan);
        var width = 0UL;
        for (var column = gridColumn; column < end; column++) width = checked(width + formatting.ColumnWidthsDxa[(int)column]);
        return Dxa(new W.TableCellWidth(), checked((uint)width));
    }

    internal static W.Shading? BuildHeaderShading(DocumentTableFormatting? formatting, bool header) =>
        formatting is not null && header
            ? new W.Shading { Val = W.ShadingPatternValues.Clear, Color = "auto", Fill = formatting.HeaderFill }
            : null;

    internal static W.RunProperties? BuildHeaderRunProperties(DocumentTableFormatting? formatting, bool header) =>
        formatting is not null && header ? new W.RunProperties(new W.Bold()) : null;

    internal static DocumentTableFormatting? Read(W.Table table, DocumentTable artifact)
    {
        var properties = table.GetFirstChild<W.TableProperties>();
        if (properties is null || properties.ChildElements.Any(child => child is not W.TableStyle and
                not W.TableWidth and not W.TableIndentation and not W.TableBorders and
                not W.TableLayout and not W.TableCellMarginDefault)) return null;
        if (!TryDxa(properties.GetFirstChild<W.TableWidth>(), positive: true, out var width) ||
            !TryDxa(properties.GetFirstChild<W.TableIndentation>(), positive: false, out var indent) ||
            properties.GetFirstChild<W.TableLayout>()?.Type?.Value != W.TableLayoutValues.Fixed)
            return null;

        var grid = table.GetFirstChild<W.TableGrid>();
        var gridColumns = grid?.Elements<W.GridColumn>().ToArray() ?? [];
        if (grid is null || grid.ChildElements.Any(child => child is not W.GridColumn) ||
            gridColumns.Length != LogicalColumns(artifact)) return null;
        var widths = new List<uint>(gridColumns.Length);
        foreach (var gridColumn in gridColumns)
        {
            if (!TryWordUInt(gridColumn, "w", positive: true, out var value)) return null;
            widths.Add(value);
        }
        if (widths.Aggregate(0UL, (sum, value) => checked(sum + value)) != width) return null;

        var margins = ReadMargins(properties.GetFirstChild<W.TableCellMarginDefault>());
        var border = ReadBorders(properties.GetFirstChild<W.TableBorders>());
        if (margins is null || border is null) return null;

        var rows = table.Elements<W.TableRow>().ToArray();
        if (rows.Length == 0 || rows.Length != artifact.Rows.Count) return null;
        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            var cells = rows[rowIndex].Elements<W.TableCell>().ToArray();
            if (cells.Length != artifact.Rows[rowIndex].RichCells.Count) return null;
            for (var cellIndex = 0; cellIndex < cells.Length; cellIndex++)
            {
                var cell = cells[cellIndex];
                var geometry = artifact.Rows[rowIndex].RichCells[cellIndex];
                var expectedWidth = widths.Skip((int)geometry.GridColumn).Take((int)geometry.ColumnSpan)
                    .Aggregate(0UL, (sum, value) => checked(sum + value));
                var cellProperties = cell.TableCellProperties;
                if (cellProperties is null || cellProperties.ChildElements.Any(child => child is not W.TableCellWidth and
                        not W.GridSpan and not W.VerticalMerge and not W.Shading) ||
                    !TryDxa(cellProperties.GetFirstChild<W.TableCellWidth>(), positive: true, out var cellWidth) ||
                    cellWidth != expectedWidth)
                    return null;

                var shading = cellProperties.GetFirstChild<W.Shading>();
                if (rowIndex == 0)
                {
                    if (shading?.Val?.Value != W.ShadingPatternValues.Clear ||
                        !string.Equals(shading.Color?.Value, "auto", StringComparison.OrdinalIgnoreCase) ||
                        !IsRgb(shading.Fill?.Value)) return null;
                    foreach (var run in cell.Descendants<W.Run>())
                    {
                        var runProperties = run.RunProperties;
                        if (runProperties is null || runProperties.ChildElements.Count != 1 ||
                            runProperties.GetFirstChild<W.Bold>() is null) return null;
                    }
                }
                else if (shading is not null || cell.Descendants<W.Bold>().Any()) return null;
            }
        }

        var headerFill = rows[0].Elements<W.TableCell>().First().TableCellProperties!.GetFirstChild<W.Shading>()!.Fill!.Value!;
        if (rows[0].Elements<W.TableCell>().Any(cell =>
                !string.Equals(cell.TableCellProperties?.GetFirstChild<W.Shading>()?.Fill?.Value, headerFill, StringComparison.OrdinalIgnoreCase)))
            return null;

        var formatting = new DocumentTableFormatting
        {
            WidthDxa = width,
            IndentDxa = indent,
            CellMarginsDxa = margins,
            BorderColor = border.Value.Color,
            BorderSize = border.Value.Size,
            HeaderFill = headerFill.ToUpperInvariant(),
        };
        formatting.ColumnWidthsDxa.Add(widths);
        return formatting;
    }

    internal static bool Same(DocumentTableFormatting? left, DocumentTableFormatting? right)
    {
        if (left is null || right is null) return left is null && right is null;
        return left.WidthDxa == right.WidthDxa && left.IndentDxa == right.IndentDxa &&
               left.ColumnWidthsDxa.SequenceEqual(right.ColumnWidthsDxa) &&
               left.CellMarginsDxa?.Top == right.CellMarginsDxa?.Top &&
               left.CellMarginsDxa?.Bottom == right.CellMarginsDxa?.Bottom &&
               left.CellMarginsDxa?.Start == right.CellMarginsDxa?.Start &&
               left.CellMarginsDxa?.End == right.CellMarginsDxa?.End &&
               left.BorderColor == right.BorderColor && left.BorderSize == right.BorderSize &&
               left.HeaderFill == right.HeaderFill;
    }

    internal static void Apply(W.Table table, DocumentTable artifact, DocumentTableFormatting formatting)
    {
        var properties = table.GetFirstChild<W.TableProperties>()!;
        SetDxa(properties.GetFirstChild<W.TableWidth>()!, formatting.WidthDxa);
        SetDxa(properties.GetFirstChild<W.TableIndentation>()!, formatting.IndentDxa);

        var borders = properties.GetFirstChild<W.TableBorders>()!;
        foreach (var border in borders.ChildElements.Cast<W.BorderType>())
        {
            border.Val = formatting.BorderSize == 0 ? W.BorderValues.Nil : W.BorderValues.Single;
            border.Size = formatting.BorderSize;
            border.Space = 0;
            border.Color = formatting.BorderColor;
        }

        var margins = properties.GetFirstChild<W.TableCellMarginDefault>()!;
        SetDxa(margins.GetFirstChild<W.TopMargin>()!, formatting.CellMarginsDxa.Top);
        SetDxa(margins.GetFirstChild<W.BottomMargin>()!, formatting.CellMarginsDxa.Bottom);
        SetDxa((OpenXmlElement?)margins.GetFirstChild<W.StartMargin>() ?? margins.GetFirstChild<W.TableCellLeftMargin>()!, formatting.CellMarginsDxa.Start);
        SetDxa((OpenXmlElement?)margins.GetFirstChild<W.EndMargin>() ?? margins.GetFirstChild<W.TableCellRightMargin>()!, formatting.CellMarginsDxa.End);

        var gridColumns = table.GetFirstChild<W.TableGrid>()!.Elements<W.GridColumn>().ToArray();
        for (var column = 0; column < gridColumns.Length; column++)
            SetWordAttribute(gridColumns[column], "w", formatting.ColumnWidthsDxa[column].ToString(CultureInfo.InvariantCulture));

        var rows = table.Elements<W.TableRow>().ToArray();
        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            var cells = rows[rowIndex].Elements<W.TableCell>().ToArray();
            for (var cellIndex = 0; cellIndex < cells.Length; cellIndex++)
            {
                var geometry = artifact.Rows[rowIndex].RichCells[cellIndex];
                var width = formatting.ColumnWidthsDxa
                    .Skip((int)geometry.GridColumn)
                    .Take((int)geometry.ColumnSpan)
                    .Aggregate(0UL, (sum, value) => checked(sum + value));
                SetDxa(cells[cellIndex].TableCellProperties!.GetFirstChild<W.TableCellWidth>()!, checked((uint)width));
                if (rowIndex == 0)
                {
                    var shading = cells[cellIndex].TableCellProperties!.GetFirstChild<W.Shading>()!;
                    shading.Val = W.ShadingPatternValues.Clear;
                    shading.Color = "auto";
                    shading.Fill = formatting.HeaderFill;
                }
            }
        }
    }

    internal static void MaskModeled(W.Table table, DocumentTable artifact)
    {
        const uint canonicalColumnWidth = 1_000;
        var columns = LogicalColumns(artifact);
        var formatting = new DocumentTableFormatting
        {
            WidthDxa = checked((uint)columns * canonicalColumnWidth),
            IndentDxa = 0,
            CellMarginsDxa = new DocumentTableCellMargins(),
            BorderColor = "000000",
            BorderSize = 2,
            HeaderFill = "000000",
        };
        for (var column = 0; column < columns; column++)
            formatting.ColumnWidthsDxa.Add(canonicalColumnWidth);
        Apply(table, artifact, formatting);
    }

    private static int LogicalColumns(DocumentTable table) => table.Rows.Any(row => row.RichCells.Count > 0)
        ? checked((int)table.GridColumns)
        : Math.Max(1, table.Rows.Count == 0 ? 1 : table.Rows.Max(row => row.Cells.Count));

    private static DocumentTableCellMargins? ReadMargins(W.TableCellMarginDefault? margins)
    {
        if (margins is null || margins.ChildElements.Count != 4 || margins.ChildElements.Any(child =>
                child is not W.TopMargin and not W.BottomMargin and not W.StartMargin and not W.EndMargin and
                not W.TableCellLeftMargin and not W.TableCellRightMargin)) return null;
        var start = (OpenXmlElement?)margins.GetFirstChild<W.StartMargin>() ?? margins.GetFirstChild<W.TableCellLeftMargin>();
        var end = (OpenXmlElement?)margins.GetFirstChild<W.EndMargin>() ?? margins.GetFirstChild<W.TableCellRightMargin>();
        if (!TryDxa(margins.GetFirstChild<W.TopMargin>(), positive: false, out var top) ||
            !TryDxa(margins.GetFirstChild<W.BottomMargin>(), positive: false, out var bottom) ||
            !TryDxa(start, positive: false, out var startValue) ||
            !TryDxa(end, positive: false, out var endValue)) return null;
        return new DocumentTableCellMargins { Top = top, Bottom = bottom, Start = startValue, End = endValue };
    }

    private static (string Color, uint Size)? ReadBorders(W.TableBorders? borders)
    {
        if (borders is null || borders.ChildElements.Count != 6 || borders.ChildElements.Any(child =>
                child is not W.TopBorder and not W.LeftBorder and not W.BottomBorder and not W.RightBorder and
                not W.InsideHorizontalBorder and not W.InsideVerticalBorder)) return null;
        string? color = null;
        uint? size = null;
        foreach (var border in borders.ChildElements)
        {
            var val = WordAttribute(border, "val");
            var currentColor = WordAttribute(border, "color").ToUpperInvariant();
            if (!TryWordUInt(border, "sz", positive: false, out var currentSize) || !IsRgb(currentColor) ||
                (currentSize == 0 ? val != "nil" : val != "single") || currentSize is 1 or > 96)
                return null;
            color ??= currentColor;
            size ??= currentSize;
            if (color != currentColor || size != currentSize) return null;
        }
        return (color!, size!.Value);
    }

    private static T Dxa<T>(T element, uint width) where T : OpenXmlElement
    {
        SetDxa(element, width);
        return element;
    }

    private static void SetDxa(OpenXmlElement element, uint width)
    {
        SetWordAttribute(element, "w", width.ToString(CultureInfo.InvariantCulture));
        SetWordAttribute(element, "type", "dxa");
    }

    private static T Border<T>(T border, DocumentTableFormatting formatting) where T : W.BorderType
    {
        border.Val = formatting.BorderSize == 0 ? W.BorderValues.Nil : W.BorderValues.Single;
        border.Size = formatting.BorderSize;
        border.Space = 0;
        border.Color = formatting.BorderColor;
        return border;
    }

    private static bool TryDxa(OpenXmlElement? element, bool positive, out uint value)
    {
        value = 0;
        return element is not null && WordAttribute(element, "type") == "dxa" &&
               TryWordUInt(element, "w", positive, out value) && value <= MaxDxa;
    }

    private static bool TryWordUInt(OpenXmlElement element, string name, bool positive, out uint value)
    {
        var parsed = uint.TryParse(WordAttribute(element, name), NumberStyles.None, CultureInfo.InvariantCulture, out value);
        return parsed && (!positive || value > 0);
    }

    private static string WordAttribute(OpenXmlElement element, string name) =>
        element.GetAttribute(name, WordprocessingNamespace).Value ?? string.Empty;

    private static void SetWordAttribute(OpenXmlElement element, string name, string value) =>
        element.SetAttribute(new OpenXmlAttribute("w", name, WordprocessingNamespace, value));

    private static bool IsRgb(string? value) => value is { Length: 6 } && value.All(character =>
        character is >= '0' and <= '9' or >= 'A' and <= 'F');

    private static CodecException Invalid(string message) => new("invalid_document_table", message);
}
