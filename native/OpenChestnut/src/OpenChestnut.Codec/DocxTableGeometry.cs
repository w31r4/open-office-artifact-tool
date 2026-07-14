using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Reads the physical WordprocessingML cell sequence onto the logical tblGrid.
// Geometry stays source-bound: the public wire exposes enough information for
// agents to understand horizontal and vertical merges without asking the
// source-preserving writer to rebuild table topology.
internal static class DocxTableGeometry
{
    private const int MaxGridColumns = 4_096;

    internal static DocumentTable Read(W.Table table, out bool valid)
    {
        var artifact = new DocumentTable();
        var declaredColumns = table.GetFirstChild<W.TableGrid>()?.Elements<W.GridColumn>().Count() ?? 0;
        valid = declaredColumns is > 0 and <= MaxGridColumns;
        var effectiveColumns = declaredColumns;
        var activeMerges = new Dictionary<(int Column, int Span), DocumentTableCell>();

        foreach (var row in table.Elements<W.TableRow>())
        {
            var targetRow = new DocumentTableRow();
            var gridBefore = NonNegative(row.TableRowProperties?.GetFirstChild<W.GridBefore>()?.Val?.Value, ref valid);
            var gridAfter = NonNegative(row.TableRowProperties?.GetFirstChild<W.GridAfter>()?.Val?.Value, ref valid);
            targetRow.GridBefore = checked((uint)gridBefore);
            targetRow.GridAfter = checked((uint)gridAfter);
            var gridColumn = gridBefore;
            var continuedMerges = new Dictionary<(int Column, int Span), DocumentTableCell>();

            foreach (var cell in row.Elements<W.TableCell>())
            {
                var span = Positive(cell.TableCellProperties?.GetFirstChild<W.GridSpan>()?.Val?.Value ?? 1, ref valid);
                var merge = VerticalMerge(cell.TableCellProperties?.GetFirstChild<W.VerticalMerge>(), ref valid);
                var text = string.Concat(cell.Descendants<W.Text>().Select(item => item.Text));
                var richCell = new DocumentTableCell
                {
                    GridColumn = checked((uint)gridColumn),
                    ColumnSpan = checked((uint)span),
                    RowSpan = merge == DocumentTableVerticalMerge.Continue ? 0u : 1u,
                    VerticalMerge = merge,
                    Editable = merge != DocumentTableVerticalMerge.Continue && IsSimpleCell(cell),
                };
                targetRow.Cells.Add(text);
                targetRow.RichCells.Add(richCell);

                var key = (gridColumn, span);
                if (merge == DocumentTableVerticalMerge.Continue)
                {
                    if (!activeMerges.TryGetValue(key, out var origin))
                    {
                        valid = false;
                        richCell.Editable = false;
                    }
                    else
                    {
                        origin.RowSpan = checked(origin.RowSpan + 1);
                        continuedMerges[key] = origin;
                    }
                }
                else if (merge == DocumentTableVerticalMerge.Restart)
                {
                    continuedMerges[key] = richCell;
                }

                gridColumn = checked(gridColumn + span);
                if (gridColumn > MaxGridColumns) valid = false;
            }

            effectiveColumns = Math.Max(effectiveColumns, checked(gridColumn + gridAfter));
            if (effectiveColumns > MaxGridColumns) valid = false;
            artifact.Rows.Add(targetRow);
            activeMerges = continuedMerges;
        }

        artifact.GridColumns = checked((uint)Math.Min(effectiveColumns, MaxGridColumns));
        return artifact;
    }

    internal static bool SameTopology(DocumentTable requested, DocumentTable source)
    {
        if (requested.GridColumns != source.GridColumns || requested.Rows.Count != source.Rows.Count) return false;
        for (var rowIndex = 0; rowIndex < source.Rows.Count; rowIndex++)
        {
            var left = requested.Rows[rowIndex];
            var right = source.Rows[rowIndex];
            if (left.GridBefore != right.GridBefore || left.GridAfter != right.GridAfter ||
                left.Cells.Count != right.Cells.Count || left.RichCells.Count != right.RichCells.Count)
                return false;
            for (var cellIndex = 0; cellIndex < right.RichCells.Count; cellIndex++)
            {
                var leftCell = left.RichCells[cellIndex];
                var rightCell = right.RichCells[cellIndex];
                if (leftCell.GridColumn != rightCell.GridColumn ||
                    leftCell.ColumnSpan != rightCell.ColumnSpan ||
                    leftCell.RowSpan != rightCell.RowSpan ||
                    leftCell.VerticalMerge != rightCell.VerticalMerge ||
                    leftCell.Editable != rightCell.Editable)
                    return false;
            }
        }
        return true;
    }

    private static int NonNegative(int? value, ref bool valid)
    {
        var resolved = value ?? 0;
        if (resolved is >= 0 and <= MaxGridColumns) return resolved;
        valid = false;
        return 0;
    }

    private static int Positive(int value, ref bool valid)
    {
        if (value is > 0 and <= MaxGridColumns) return value;
        valid = false;
        return 1;
    }

    private static DocumentTableVerticalMerge VerticalMerge(W.VerticalMerge? merge, ref bool valid)
    {
        if (merge is null) return DocumentTableVerticalMerge.Unspecified;
        var value = merge.Val?.Value;
        if (value is null || value == W.MergedCellValues.Continue) return DocumentTableVerticalMerge.Continue;
        if (value == W.MergedCellValues.Restart) return DocumentTableVerticalMerge.Restart;
        valid = false;
        return DocumentTableVerticalMerge.Unspecified;
    }

    internal static bool IsSimpleCell(W.TableCell cell)
    {
        if (cell.ChildElements.Any(child => child is not W.TableCellProperties and not W.Paragraph)) return false;
        var paragraphs = cell.Elements<W.Paragraph>().ToArray();
        if (paragraphs.Length != 1) return false;
        var paragraph = paragraphs[0];
        if (paragraph.ChildElements.Any(child => child is not W.ParagraphProperties and not W.Run)) return false;
        var runs = paragraph.Elements<W.Run>().ToArray();
        if (runs.Length != 1) return false;
        var run = runs[0];
        return !run.ChildElements.Any(child => child is not W.RunProperties and not W.Text) &&
               run.Elements<W.Text>().Count() == 1;
    }
}
