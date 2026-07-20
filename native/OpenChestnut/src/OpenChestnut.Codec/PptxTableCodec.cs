using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenChestnut.Codec;

// Owns one deliberately narrow DrawingML table profile. Table topology,
// rectangular merge ranges, and cell formatting remain fixed after import;
// name, complete outer frame, and the single plain-text run in each visible
// origin cell are the only source-bound edits.
internal static class PptxTableCodec
{
    private const string TableGraphicDataUri = "http://schemas.openxmlformats.org/drawingml/2006/table";
    private const int MaxColumns = 256;
    private const int MaxRows = 2_048;
    private const int MaxCellTextLength = 32_767;

    private readonly record struct MergeCellPlan(
        bool IsOrigin,
        int RowSpan,
        int ColumnSpan,
        bool HorizontalMerge,
        bool VerticalMerge);

    private readonly record struct NativeMergeCell(
        int RowSpan,
        int ColumnSpan,
        bool HorizontalMerge,
        bool VerticalMerge,
        bool HasRowSpan,
        bool HasColumnSpan,
        bool HasHorizontalMerge,
        bool HasVerticalMerge);

    internal static bool TryRead(P.GraphicFrame source, out PresentationTable table)
    {
        table = new PresentationTable();
        try
        {
            if (source.ChildElements.Count != 3 ||
                source.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties is not { Id.HasValue: true, Name: not null } ||
                source.Transform is not { } transform ||
                !TryReadFrame(transform, out var left, out var top, out var width, out var height) ||
                source.Graphic is not { ChildElements.Count: 1 } graphic ||
                graphic.GraphicData is not { ChildElements.Count: 1 } graphicData ||
                !string.Equals(graphicData.Uri?.Value, TableGraphicDataUri, StringComparison.Ordinal) ||
                graphicData.GetFirstChild<A.Table>() is not { } nativeTable ||
                nativeTable.ChildElements.Any(child => child is not A.TableProperties and not A.TableGrid and not A.TableRow))
                return false;

            var properties = nativeTable.Elements<A.TableProperties>().SingleOrDefault();
            var grid = nativeTable.Elements<A.TableGrid>().SingleOrDefault();
            var rows = nativeTable.Elements<A.TableRow>().ToArray();
            if (properties is null || grid is null ||
                nativeTable.Elements<A.TableProperties>().Count() != 1 ||
                nativeTable.Elements<A.TableGrid>().Count() != 1 ||
                nativeTable.ChildElements[0] is not A.TableProperties ||
                nativeTable.ChildElements[1] is not A.TableGrid ||
                properties.ChildElements.Count != 0 ||
                !HasOnlyAttributes(properties, "firstRow", "bandRow") ||
                rows.Length is < 1 or > MaxRows)
                return false;

            var columns = grid.Elements<A.GridColumn>().ToArray();
            if (columns.Length is < 1 or > MaxColumns ||
                grid.ChildElements.Count != columns.Length ||
                columns.Any(column => column.Width?.Value is null or <= 0 || column.ChildElements.Count != 0 || !HasOnlyAttributes(column, "w")))
                return false;

            var result = new PresentationTable
            {
                LeftEmu = left,
                TopEmu = top,
                WidthEmu = width,
                HeightEmu = height,
            };
            result.ColumnWidthsEmu.Add(columns.Select(column => column.Width!.Value));
            if (properties.FirstRow is not null) result.FirstRow = properties.FirstRow.Value;
            if (properties.BandRow is not null) result.BandedRows = properties.BandRow.Value;

            var nativeCells = new List<A.TableCell[]>(rows.Length);
            foreach (var nativeRow in rows)
            {
                if (nativeRow.Height?.Value is null or <= 0 || !HasOnlyAttributes(nativeRow, "h")) return false;
                var cells = nativeRow.Elements<A.TableCell>().ToArray();
                if (nativeRow.ChildElements.Count != cells.Length || cells.Length != columns.Length) return false;
                nativeCells.Add(cells);
                var row = new PresentationTableRow { HeightEmu = nativeRow.Height.Value };
                foreach (var cell in cells)
                {
                    if (!TryReadCell(cell, out var text)) return false;
                    row.Cells.Add(new PresentationTableCell { Text = text.Text });
                }
                result.Rows.Add(row);
            }

            if (!TryReadMergeRanges(nativeCells, result)) return false;
            if (result.ColumnWidthsEmu.Sum() != width || result.Rows.Sum(row => row.HeightEmu) != height) return false;
            table = result;
            return true;
        }
        catch (Exception error) when (error is InvalidOperationException or OverflowException)
        {
            table = new PresentationTable();
            return false;
        }
    }

    internal static P.GraphicFrame Build(PresentationElement element, uint nativeId)
    {
        var table = element.Table;
        Validate(table, element.Id);
        var mergePlan = CreateMergePlan(table, element.Id);
        var properties = new A.TableProperties();
        if (table.HasFirstRow) properties.FirstRow = table.FirstRow;
        if (table.HasBandedRows) properties.BandRow = table.BandedRows;
        var grid = new A.TableGrid();
        foreach (var width in table.ColumnWidthsEmu) grid.Append(new A.GridColumn { Width = width });
        var nativeTable = new A.Table(properties, grid);
        for (var rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            var sourceRow = table.Rows[rowIndex];
            var row = new A.TableRow { Height = sourceRow.HeightEmu };
            for (var columnIndex = 0; columnIndex < sourceRow.Cells.Count; columnIndex++)
            {
                var sourceCell = sourceRow.Cells[columnIndex];
                var cell = BuildCell(sourceCell.Text, table.HasFirstRow && table.FirstRow && rowIndex == 0);
                if (mergePlan.TryGetValue((rowIndex, columnIndex), out var merge))
                {
                    if (merge.IsOrigin)
                    {
                        if (merge.RowSpan > 1) cell.RowSpan = merge.RowSpan;
                        if (merge.ColumnSpan > 1) cell.GridSpan = merge.ColumnSpan;
                    }
                    else
                    {
                        if (merge.HorizontalMerge) cell.HorizontalMerge = true;
                        if (merge.VerticalMerge) cell.VerticalMerge = true;
                    }
                }
                row.Append(cell);
            }
            nativeTable.Append(row);
        }
        return new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = nativeId, Name = element.Name },
                new P.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.Transform(
                new A.Offset { X = table.LeftEmu, Y = table.TopEmu },
                new A.Extents { Cx = table.WidthEmu, Cy = table.HeightEmu }),
            new A.Graphic(new A.GraphicData(nativeTable) { Uri = TableGraphicDataUri }));
    }

    internal static void Apply(P.GraphicFrame source, PresentationElement requested)
    {
        if (!TryRead(source, out var original))
            throw new CodecException("unsupported_presentation_edit", $"Presentation table {requested.Id} no longer matches the editable table profile.");
        ValidateRequest(original, requested);
        var table = requested.Table;
        source.NonVisualGraphicFrameProperties!.NonVisualDrawingProperties!.Name = requested.Name;
        SetFrame(source.Transform!, table);
        var nativeTable = source.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
        var columns = nativeTable.GetFirstChild<A.TableGrid>()!.Elements<A.GridColumn>().ToArray();
        for (var index = 0; index < columns.Length; index++) columns[index].Width = table.ColumnWidthsEmu[index];
        var rows = nativeTable.Elements<A.TableRow>().ToArray();
        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            rows[rowIndex].Height = table.Rows[rowIndex].HeightEmu;
            var cells = rows[rowIndex].Elements<A.TableCell>().ToArray();
            for (var columnIndex = 0; columnIndex < cells.Length; columnIndex++)
                SingleText(cells[columnIndex]).Text = table.Rows[rowIndex].Cells[columnIndex].Text;
        }
    }

    internal static void Validate(PresentationTable? table, string elementId)
    {
        if (table is null) throw Invalid(elementId, "payload is missing");
        if (table.LeftEmu < 0 || table.TopEmu < 0 || table.WidthEmu <= 0 || table.HeightEmu <= 0)
            throw Invalid(elementId, "frame must have non-negative coordinates and positive dimensions");
        if (table.ColumnWidthsEmu.Count is < 1 or > MaxColumns || table.Rows.Count is < 1 or > MaxRows)
            throw Invalid(elementId, $"grid must contain 1-{MaxColumns} columns and 1-{MaxRows} rows");
        if (table.ColumnWidthsEmu.Any(width => width <= 0) || Sum(table.ColumnWidthsEmu, elementId) != table.WidthEmu)
            throw Invalid(elementId, "positive column widths must sum to the outer frame width");
        if (table.Rows.Any(row => row.HeightEmu <= 0 || row.Cells.Count != table.ColumnWidthsEmu.Count) ||
            Sum(table.Rows.Select(row => row.HeightEmu), elementId) != table.HeightEmu)
            throw Invalid(elementId, "positive row heights must sum to the outer frame height and every row must match the grid width");
        foreach (var cell in table.Rows.SelectMany(row => row.Cells))
            if (cell.Text.Length > MaxCellTextLength || cell.Text.Any(character => char.IsControl(character) && character is not '\t' and not '\n' and not '\r'))
                throw Invalid(elementId, $"cell text must contain at most {MaxCellTextLength} characters and no unsupported controls");
        _ = CreateMergePlan(table, elementId);
    }

    internal static void ScrubModeledContent(P.GraphicFrame source)
    {
        if (source.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties is { } nonVisual) nonVisual.Name = string.Empty;
        if (source.Transform is { } transform)
        {
            transform.Offset!.X = 0L;
            transform.Offset.Y = 0L;
            transform.Extents!.Cx = 1L;
            transform.Extents.Cy = 1L;
        }
        var table = source.Graphic?.GraphicData?.GetFirstChild<A.Table>();
        if (table is null) return;
        foreach (var column in table.GetFirstChild<A.TableGrid>()?.Elements<A.GridColumn>() ?? []) column.Width = 1L;
        foreach (var row in table.Elements<A.TableRow>())
        {
            row.Height = 1L;
            foreach (var cell in row.Elements<A.TableCell>()) SingleText(cell).Text = string.Empty;
        }
    }

    private static void ValidateRequest(PresentationTable original, PresentationElement requested)
    {
        Validate(requested.Table, requested.Id);
        if (requested.Name.Length > 1_024) throw Invalid(requested.Id, "name exceeds 1024 characters");
        var allowed = original.Clone();
        allowed.LeftEmu = requested.Table.LeftEmu;
        allowed.TopEmu = requested.Table.TopEmu;
        allowed.WidthEmu = requested.Table.WidthEmu;
        allowed.HeightEmu = requested.Table.HeightEmu;
        allowed.ColumnWidthsEmu.Clear();
        allowed.ColumnWidthsEmu.Add(requested.Table.ColumnWidthsEmu);
        for (var rowIndex = 0; rowIndex < allowed.Rows.Count; rowIndex++)
        {
            allowed.Rows[rowIndex].HeightEmu = requested.Table.Rows[rowIndex].HeightEmu;
            for (var columnIndex = 0; columnIndex < allowed.Rows[rowIndex].Cells.Count; columnIndex++)
                allowed.Rows[rowIndex].Cells[columnIndex].Text = requested.Table.Rows[rowIndex].Cells[columnIndex].Text;
        }
        if (!allowed.Equals(requested.Table))
            throw new CodecException("unsupported_presentation_edit", $"Presentation table {requested.Id} may edit only its name, complete frame, and fixed-topology plain cell text.");
    }

    private static bool TryReadFrame(P.Transform transform, out long left, out long top, out long width, out long height)
    {
        left = top = width = height = 0;
        if (transform.ChildElements.Count != 2 || transform.Offset is not { } offset || transform.Extents is not { } extents ||
            !HasOnlyAttributes(offset, "x", "y") || !HasOnlyAttributes(extents, "cx", "cy") ||
            offset.X?.Value is null || offset.Y?.Value is null || extents.Cx?.Value is null or <= 0 || extents.Cy?.Value is null or <= 0 ||
            offset.X.Value < 0 || offset.Y.Value < 0)
            return false;
        left = offset.X.Value;
        top = offset.Y.Value;
        width = extents.Cx.Value;
        height = extents.Cy.Value;
        return true;
    }

    private static bool TryReadCell(A.TableCell cell, out A.Text text)
    {
        text = new A.Text();
        if (!HasOnlyAttributes(cell, "rowSpan", "gridSpan", "hMerge", "vMerge") || cell.ChildElements.Count != 2 ||
            cell.ChildElements[0] is not A.TextBody body || cell.ChildElements[1] is not A.TableCellProperties ||
            body.ChildElements.Count != 3 || body.ChildElements[0] is not A.BodyProperties || body.ChildElements[1] is not A.ListStyle ||
            body.ChildElements[2] is not A.Paragraph paragraph)
            return false;
        var run = paragraph.Elements<A.Run>().SingleOrDefault();
        if (run is null || paragraph.Elements<A.Run>().Count() != 1 ||
            paragraph.ChildElements.Any(child => child is not A.Run and not A.EndParagraphRunProperties) ||
            paragraph.Elements<A.EndParagraphRunProperties>().Count() > 1 ||
            run.ChildElements.Any(child => child is not A.RunProperties and not A.Text) ||
            run.Elements<A.RunProperties>().Count() > 1 || run.Elements<A.Text>().Count() != 1)
            return false;
        text = run.GetFirstChild<A.Text>()!;
        return text.Text.Length <= MaxCellTextLength;
    }

    private static bool TryReadMergeRanges(IReadOnlyList<A.TableCell[]> nativeRows, PresentationTable table)
    {
        var rowCount = nativeRows.Count;
        var columnCount = nativeRows[0].Length;
        var cells = new NativeMergeCell[rowCount, columnCount];
        for (var row = 0; row < rowCount; row++)
        {
            for (var column = 0; column < columnCount; column++)
            {
                var cell = nativeRows[row][column];
                var hasRowSpan = cell.RowSpan is not null;
                var hasColumnSpan = cell.GridSpan is not null;
                var hasHorizontal = cell.HorizontalMerge is not null;
                var hasVertical = cell.VerticalMerge is not null;
                var rowSpan = cell.RowSpan?.Value ?? 1;
                var columnSpan = cell.GridSpan?.Value ?? 1;
                var horizontal = cell.HorizontalMerge?.Value ?? false;
                var vertical = cell.VerticalMerge?.Value ?? false;
                if ((hasRowSpan && rowSpan <= 1) || (hasColumnSpan && columnSpan <= 1) ||
                    (hasHorizontal && !horizontal) || (hasVertical && !vertical) ||
                    ((horizontal || vertical) && (hasRowSpan || hasColumnSpan)))
                    return false;
                cells[row, column] = new NativeMergeCell(rowSpan, columnSpan, horizontal, vertical, hasRowSpan, hasColumnSpan, hasHorizontal, hasVertical);
            }
        }

        var expected = new Dictionary<(int Row, int Column), MergeCellPlan>();
        for (var row = 0; row < rowCount; row++)
        {
            for (var column = 0; column < columnCount; column++)
            {
                var cell = cells[row, column];
                if (cell.HorizontalMerge || cell.VerticalMerge || cell.RowSpan == 1 && cell.ColumnSpan == 1) continue;
                if ((long)row + cell.RowSpan > rowCount || (long)column + cell.ColumnSpan > columnCount) return false;
                var range = new PresentationTableMergeRange
                {
                    StartRow = (uint)row,
                    EndRow = (uint)(row + cell.RowSpan - 1),
                    StartColumn = (uint)column,
                    EndColumn = (uint)(column + cell.ColumnSpan - 1),
                };
                for (var coveredRow = row; coveredRow <= range.EndRow; coveredRow++)
                {
                    for (var coveredColumn = column; coveredColumn <= range.EndColumn; coveredColumn++)
                    {
                        var isOrigin = coveredRow == row && coveredColumn == column;
                        if (!expected.TryAdd((coveredRow, coveredColumn), new MergeCellPlan(
                            isOrigin,
                            isOrigin ? cell.RowSpan : 0,
                            isOrigin ? cell.ColumnSpan : 0,
                            !isOrigin && coveredColumn > column,
                            !isOrigin && coveredRow > row)))
                            return false;
                    }
                }
                table.MergeRanges.Add(range);
            }
        }

        for (var row = 0; row < rowCount; row++)
        {
            for (var column = 0; column < columnCount; column++)
            {
                var cell = cells[row, column];
                if (expected.TryGetValue((row, column), out var planned))
                {
                    if (planned.IsOrigin)
                    {
                        if (cell.HorizontalMerge || cell.VerticalMerge || cell.RowSpan != planned.RowSpan || cell.ColumnSpan != planned.ColumnSpan) return false;
                    }
                    else
                    {
                        if (cell.HasRowSpan || cell.HasColumnSpan || cell.HorizontalMerge != planned.HorizontalMerge || cell.VerticalMerge != planned.VerticalMerge ||
                            !string.IsNullOrEmpty(table.Rows[row].Cells[column].Text))
                            return false;
                    }
                }
                else if (cell.HasRowSpan || cell.HasColumnSpan || cell.HasHorizontalMerge || cell.HasVerticalMerge)
                {
                    return false;
                }
            }
        }
        return true;
    }

    private static Dictionary<(int Row, int Column), MergeCellPlan> CreateMergePlan(PresentationTable table, string elementId)
    {
        var plan = new Dictionary<(int Row, int Column), MergeCellPlan>();
        for (var rangeIndex = 0; rangeIndex < table.MergeRanges.Count; rangeIndex++)
        {
            var range = table.MergeRanges[rangeIndex];
            if (range.EndRow < range.StartRow || range.EndColumn < range.StartColumn ||
                range.EndRow >= table.Rows.Count || range.EndColumn >= table.ColumnWidthsEmu.Count ||
                range.StartRow == range.EndRow && range.StartColumn == range.EndColumn)
                throw Invalid(elementId, $"merge range {rangeIndex} must cover at least two in-bounds cells");
            var rowSpan = checked((int)(range.EndRow - range.StartRow + 1));
            var columnSpan = checked((int)(range.EndColumn - range.StartColumn + 1));
            for (var row = checked((int)range.StartRow); row <= range.EndRow; row++)
            {
                for (var column = checked((int)range.StartColumn); column <= range.EndColumn; column++)
                {
                    var isOrigin = row == range.StartRow && column == range.StartColumn;
                    if (!plan.TryAdd((row, column), new MergeCellPlan(
                        isOrigin,
                        isOrigin ? rowSpan : 0,
                        isOrigin ? columnSpan : 0,
                        !isOrigin && column > range.StartColumn,
                        !isOrigin && row > range.StartRow)))
                        throw Invalid(elementId, $"merge ranges overlap at cell {row},{column}");
                    if (!isOrigin && !string.IsNullOrEmpty(table.Rows[row].Cells[column].Text))
                        throw Invalid(elementId, $"covered merge cell {row},{column} must be empty");
                }
            }
        }
        return plan;
    }

    private static A.TableCell BuildCell(string text, bool header)
    {
        var runProperties = new A.RunProperties { Language = "en-US", FontSize = 1_350, Bold = header };
        runProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = header ? "000000" : "0F172A" }));
        var cellProperties = new A.TableCellProperties(
            new A.SolidFill(new A.RgbColorModelHex { Val = header ? "EDEDED" : "FFFFFF" }));
        return new A.TableCell(
            new A.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(
                    new A.Run(runProperties, new A.Text(text)),
                    new A.EndParagraphRunProperties { Language = "en-US", FontSize = 1_350 })),
            cellProperties);
    }

    private static A.Text SingleText(A.TableCell cell) => cell.GetFirstChild<A.TextBody>()!.Descendants<A.Text>().Single();

    private static void SetFrame(P.Transform transform, PresentationTable table)
    {
        transform.Offset!.X = table.LeftEmu;
        transform.Offset.Y = table.TopEmu;
        transform.Extents!.Cx = table.WidthEmu;
        transform.Extents.Cy = table.HeightEmu;
    }

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute =>
            string.IsNullOrEmpty(attribute.NamespaceUri) && allowed.Contains(attribute.LocalName));
    }

    private static long Sum(IEnumerable<long> values, string elementId)
    {
        try { return values.Aggregate(0L, checked((total, value) => total + value)); }
        catch (OverflowException) { throw Invalid(elementId, "grid dimensions overflow the supported EMU range"); }
    }

    private static CodecException Invalid(string elementId, string message) =>
        new("invalid_presentation_table", $"Presentation table {elementId} {message}.");
}
