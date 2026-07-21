using System.Security.Cryptography;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Owns fixed-topology source edits and bounded source-free table construction.
// Imported table, row, cell, paragraph, and run properties stay in source XML
// behind the residual hash except for the exact direct-formatting profile owned
// by DocxTableFormatting. Direct geometry must completely cover tblGrid and
// carry an exact gridSpan/vMerge chain; ambiguous continuations fail closed.
internal static class DocxTableCodec
{
    internal static DocumentTable Read(W.Table table, out bool editable)
    {
        var artifact = DocxTableGeometry.Read(table, out var validGeometry);
        artifact.Formatting = DocxTableFormatting.Read(table, artifact);
        editable = validGeometry && HasSafeContainerTopology(table) &&
                   artifact.Rows.SelectMany(row => row.RichCells).Any(cell => cell.Editable || cell.TextPatchable);
        return artifact;
    }

    internal static void Apply(W.Table table, DocumentTable requested)
    {
        Validate(requested);
        var source = Read(table, out var editable);
        if (!editable) throw Unsupported("Source-preserving DOCX export cannot edit this table topology.");
        if (!DocxTableGeometry.SameTopology(requested, source))
            throw Unsupported("Source-preserving DOCX table grid, span, and merge topology cannot be changed.");
        if (!DocxTableFormatting.Same(requested.Formatting, source.Formatting))
        {
            if (source.Formatting is null || requested.Formatting is null)
                throw Unsupported("Source-preserving DOCX table formatting can change only when the complete direct-formatting profile was recognized during import.");
            DocxTableFormatting.Apply(table, requested, requested.Formatting);
            if (!DocxTableFormatting.Same(DocxTableFormatting.Read(table, requested), requested.Formatting))
                throw Unsupported("Source-preserving DOCX table formatting did not round trip through the bounded direct-formatting profile.");
        }

        var rows = table.Elements<W.TableRow>().ToArray();
        if (rows.Length != requested.Rows.Count)
            throw Unsupported("Source-preserving DOCX table row topology cannot be changed.");
        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            var cells = rows[rowIndex].Elements<W.TableCell>().ToArray();
            if (cells.Length != requested.Rows[rowIndex].Cells.Count)
                throw Unsupported($"Source-preserving DOCX table row {rowIndex} cell topology cannot be changed.");
            for (var cellIndex = 0; cellIndex < cells.Length; cellIndex++)
            {
                var value = requested.Rows[rowIndex].Cells[cellIndex];
                if (value == source.Rows[rowIndex].Cells[cellIndex]) continue;
                if (!source.Rows[rowIndex].RichCells[cellIndex].Editable)
                    throw Unsupported($"Source-preserving DOCX table cell {rowIndex},{cellIndex} is a continuation or complex cell and cannot be edited.");
                var text = cells[cellIndex].Descendants<W.Text>().Single();
                text.Text = value;
                text.Space = value.Length != value.Trim().Length ? SpaceProcessingModeValues.Preserve : null;
            }
        }
        ApplyTextPatches(rows, requested, source);
    }

    internal static W.Table Build(DocumentBlock block)
    {
        Validate(block.Table);
        if (block.Table.TextPatches.Count > 0)
            throw Unsupported("DOCX table text patches require a validated imported source package.");
        if (!string.IsNullOrWhiteSpace(block.StyleId) && !block.StyleId.Equals("TableGrid", StringComparison.Ordinal))
            throw new CodecException(
                "unsupported_document_features",
                $"Direct DOCX table authoring cannot materialize custom table style {block.StyleId} without a modeled style graph.");
        var table = new W.Table();
        var tableProperties = DocxTableFormatting.BuildProperties(block.StyleId, block.Table.Formatting);
        if (tableProperties is not null) table.Append(tableProperties);

        if (block.Table.Rows.All(row => row.RichCells.Count == 0))
        {
            var columns = block.Table.Rows.Count == 0 ? 1 : Math.Max(1, block.Table.Rows.Max(row => row.Cells.Count));
            var simpleGrid = new W.TableGrid();
            for (var column = 0; column < columns; column++)
                simpleGrid.Append(DocxTableFormatting.BuildGridColumn(block.Table.Formatting, column));
            table.Append(simpleGrid);
            for (var rowIndex = 0; rowIndex < block.Table.Rows.Count; rowIndex++)
            {
                var sourceRow = block.Table.Rows[rowIndex];
                var row = new W.TableRow();
                for (var cellIndex = 0; cellIndex < sourceRow.Cells.Count; cellIndex++)
                    row.Append(BuildCell(
                        sourceRow.Cells[cellIndex],
                        formatting: block.Table.Formatting,
                        gridColumn: checked((uint)cellIndex),
                        header: rowIndex == 0));
                table.Append(row);
            }
            return table;
        }

        ValidateAuthoredGeometry(block.Table);
        var grid = new W.TableGrid();
        for (var column = 0; column < block.Table.GridColumns; column++)
            grid.Append(DocxTableFormatting.BuildGridColumn(block.Table.Formatting, column));
        table.Append(grid);
        for (var rowIndex = 0; rowIndex < block.Table.Rows.Count; rowIndex++)
        {
            var sourceRow = block.Table.Rows[rowIndex];
            var row = new W.TableRow();
            if (sourceRow.GridBefore != 0 || sourceRow.GridAfter != 0)
            {
                var properties = new W.TableRowProperties();
                if (sourceRow.GridBefore != 0) properties.Append(new W.GridBefore { Val = checked((int)sourceRow.GridBefore) });
                if (sourceRow.GridAfter != 0) properties.Append(new W.GridAfter { Val = checked((int)sourceRow.GridAfter) });
                row.Append(properties);
            }
            for (var cellIndex = 0; cellIndex < sourceRow.Cells.Count; cellIndex++)
                row.Append(BuildCell(
                    sourceRow.Cells[cellIndex],
                    sourceRow.RichCells[cellIndex],
                    block.Table.Formatting,
                    sourceRow.RichCells[cellIndex].GridColumn,
                    rowIndex == 0));
            table.Append(row);
        }
        return table;
    }

    internal static string ResidualHash(W.Table table)
    {
        var clone = (W.Table)table.CloneNode(true);
        var artifact = DocxTableGeometry.Read(clone, out _);
        if (DocxTableFormatting.Read(clone, artifact) is not null)
            DocxTableFormatting.MaskModeled(clone, artifact);
        foreach (var text in clone.Descendants<W.Text>())
        {
            text.Text = string.Empty;
            text.Space = null;
        }
        return Hash(Encoding.UTF8.GetBytes(clone.OuterXml));
    }

    internal static void Validate(DocumentTable? table)
    {
        if (table is null) throw Invalid("Document table payload is missing.");
        if (table.TextPatches.Count > 10_000) throw Invalid("Document table exceeds 10,000 source text patches.");
        DocxTableFormatting.Validate(table);
        for (var rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            var row = table.Rows[rowIndex];
            for (var cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
            {
                if (row.Cells[cellIndex].Length > 1_000_000)
                    throw Invalid($"Document table cell {rowIndex},{cellIndex} exceeds 1,000,000 characters.");
            }
            if (row.RichCells.Count == 0)
            {
                if (row.GridBefore != 0 || row.GridAfter != 0)
                    throw Invalid($"Document table row {rowIndex} cannot declare grid offsets without rich_cells.");
                continue;
            }
            for (var cellIndex = 0; cellIndex < row.RichCells.Count; cellIndex++)
            {
                var cell = row.RichCells[cellIndex];
                if (cell.ColumnSpan is 0 or > 4_096 || cell.GridColumn > 4_096 || cell.RowSpan > 4_096)
                    throw Invalid($"Document table cell {rowIndex},{cellIndex} has invalid bounded grid geometry.");
                if (cell.VerticalMerge == DocumentTableVerticalMerge.Continue && (cell.RowSpan != 0 || cell.Editable))
                    throw Invalid($"Document table continuation cell {rowIndex},{cellIndex} must be read-only with row_span zero.");
                if (cell.VerticalMerge == DocumentTableVerticalMerge.Continue && cell.TextPatchable)
                    throw Invalid($"Document table continuation cell {rowIndex},{cellIndex} cannot be text-patchable.");
                if (cell.VerticalMerge != DocumentTableVerticalMerge.Continue && cell.RowSpan == 0)
                    throw Invalid($"Document table origin cell {rowIndex},{cellIndex} must have a positive row_span.");
            }
        }
        foreach (var patch in table.TextPatches)
        {
            if (patch.Row >= table.Rows.Count || patch.Column >= table.Rows[(int)patch.Row].Cells.Count)
                throw Invalid($"Document table text patch {patch.Row},{patch.Column} is outside the physical cell matrix.");
            if (string.IsNullOrEmpty(patch.Search) || patch.Search.Length > 1_000_000 || patch.Replacement.Length > 1_000_000 ||
                !XmlSafe(patch.Search) || !XmlSafe(patch.Replacement))
                throw Invalid($"Document table text patch {patch.Row},{patch.Column} requires bounded XML-safe search and replacement strings.");
            if (patch.SourceTextSha256.Length != 64 || patch.SourceTextSha256.Any(character => !Uri.IsHexDigit(character)))
                throw Invalid($"Document table text patch {patch.Row},{patch.Column} requires a SHA-256 source text binding.");
        }
        if (table.Rows.Any(row => row.RichCells.Count > 0) && table.GridColumns is 0 or > 4_096)
            throw Invalid("Document table grid_columns must be between 1 and 4096 when rich cell geometry is present.");
    }

    internal static bool SemanticsMatchRequested(W.Table table, DocumentTable requested)
    {
        var actual = Read(table, out _);
        var expected = requested.Clone();
        foreach (var patch in expected.TextPatches)
        {
            var value = expected.Rows[(int)patch.Row].Cells[(int)patch.Column];
            var index = value.IndexOf(patch.Search, StringComparison.Ordinal);
            if (index < 0 || value.IndexOf(patch.Search, index + 1, StringComparison.Ordinal) >= 0) return false;
            expected.Rows[(int)patch.Row].Cells[(int)patch.Column] =
                value.Remove(index, patch.Search.Length).Insert(index, patch.Replacement);
        }
        expected.TextPatches.Clear();
        return expected.Equals(actual);
    }

    private static void ApplyTextPatches(W.TableRow[] rows, DocumentTable requested, DocumentTable source)
    {
        foreach (var group in requested.TextPatches.GroupBy(patch => (patch.Row, patch.Column)))
        {
            var rowIndex = checked((int)group.Key.Row);
            var cellIndex = checked((int)group.Key.Column);
            var sourceValue = source.Rows[rowIndex].Cells[cellIndex];
            if (requested.Rows[rowIndex].Cells[cellIndex] != sourceValue)
                throw Unsupported($"Source-preserving DOCX table cell {rowIndex},{cellIndex} cannot combine whole-cell and native text-patch edits.");
            var sourceCell = source.Rows[rowIndex].RichCells[cellIndex];
            if (!sourceCell.TextPatchable)
                throw Unsupported($"Source-preserving DOCX table cell {rowIndex},{cellIndex} contains no safely patchable plain text node.");
            var sourceHash = Hash(Encoding.UTF8.GetBytes(sourceValue));
            var cell = rows[rowIndex].Elements<W.TableCell>().ElementAt(cellIndex);
            var expected = sourceValue;
            foreach (var patch in group)
            {
                if (!sourceHash.Equals(patch.SourceTextSha256, StringComparison.OrdinalIgnoreCase))
                    throw Unsupported($"Source-preserving DOCX table cell {rowIndex},{cellIndex} text no longer matches the patch source binding.");
                var offset = expected.IndexOf(patch.Search, StringComparison.Ordinal);
                if (offset < 0 || expected.IndexOf(patch.Search, offset + 1, StringComparison.Ordinal) >= 0)
                    throw Unsupported($"Source-preserving DOCX table cell {rowIndex},{cellIndex} text patch requires exactly one visible match.");
                var resolution = DocxLiteralTextSpanCodec.Resolve(cell, expected, patch.Search);
                if (resolution.Status == DocxLiteralTextSpanStatus.TextMismatch)
                    throw Unsupported($"Source-preserving DOCX table cell {rowIndex},{cellIndex} native text no longer matches its semantic source snapshot.");
                if (resolution.Status == DocxLiteralTextSpanStatus.MatchNotUnique)
                    throw Unsupported($"Source-preserving DOCX table cell {rowIndex},{cellIndex} text patch requires exactly one visible match; found {resolution.MatchCount}.");
                if (resolution.Status != DocxLiteralTextSpanStatus.Success || resolution.Span is null)
                    throw Unsupported($"Source-preserving DOCX table cell {rowIndex},{cellIndex} text patch must stay inside one ordinary native text node or adjacent same-format runs; {DocxLiteralTextSpanCodec.FailureDescription(resolution.Status)}.");
                expected = expected.Remove(offset, patch.Search.Length).Insert(offset, patch.Replacement);
                DocxLiteralTextSpanCodec.Replace(resolution.Span, patch.Replacement);
            }
            if (!string.Concat(cell.Descendants<W.Text>().Select(text => text.Text)).Equals(expected, StringComparison.Ordinal))
                throw new CodecException(
                    "document_semantics_not_applied",
                    $"Source-preserving DOCX table cell {rowIndex},{cellIndex} text patches did not produce the requested visible text.",
                    "word/document.xml");
        }
    }

    private static bool XmlSafe(string value)
    {
        try
        {
            XmlConvert.VerifyXmlChars(value);
            return !value.Contains('\u007f');
        }
        catch (XmlException)
        {
            return false;
        }
    }

    private static void ValidateAuthoredGeometry(DocumentTable table)
    {
        if (table.Rows.Count == 0) throw Invalid("Authored document table geometry requires at least one row.");
        var active = new Dictionary<(uint Column, uint Span), MergeGroup>();
        for (var rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            var row = table.Rows[rowIndex];
            if (row.Cells.Count == 0 || row.RichCells.Count != row.Cells.Count)
                throw Invalid($"Authored document table row {rowIndex} requires one geometry record for every physical cell.");
            if (row.GridBefore > table.GridColumns || row.GridAfter > table.GridColumns)
                throw Invalid($"Authored document table row {rowIndex} has a grid offset outside grid_columns.");

            var cursor = row.GridBefore;
            var continued = new Dictionary<(uint Column, uint Span), MergeGroup>();
            for (var cellIndex = 0; cellIndex < row.RichCells.Count; cellIndex++)
            {
                var cell = row.RichCells[cellIndex];
                if (cell.GridColumn != cursor)
                    throw Invalid($"Authored document table cell {rowIndex},{cellIndex} must begin at grid column {cursor}.");
                cursor = checked(cursor + cell.ColumnSpan);
                if (cursor > table.GridColumns)
                    throw Invalid($"Authored document table cell {rowIndex},{cellIndex} extends beyond grid_columns.");
                var key = (cell.GridColumn, cell.ColumnSpan);
                switch (cell.VerticalMerge)
                {
                    case DocumentTableVerticalMerge.Unspecified:
                        if (cell.RowSpan != 1 || !cell.Editable)
                            throw Invalid($"Authored unmerged cell {rowIndex},{cellIndex} must have row_span one and remain editable.");
                        break;
                    case DocumentTableVerticalMerge.Restart:
                        if (cell.RowSpan == 0 || !cell.Editable)
                            throw Invalid($"Authored merge origin {rowIndex},{cellIndex} must have a positive row_span and remain editable.");
                        continued.Add(key, new MergeGroup(rowIndex, cellIndex, checked((int)cell.RowSpan)));
                        break;
                    case DocumentTableVerticalMerge.Continue:
                        if (cell.RowSpan != 0 || cell.Editable || row.Cells[cellIndex].Length != 0)
                            throw Invalid($"Authored merge continuation {rowIndex},{cellIndex} must be read-only with row_span zero and empty text.");
                        if (!active.TryGetValue(key, out var group))
                            throw Invalid($"Authored merge continuation {rowIndex},{cellIndex} has no matching restart in the preceding row.");
                        group.Seen++;
                        continued.Add(key, group);
                        break;
                    default:
                        throw Invalid($"Authored document table cell {rowIndex},{cellIndex} has an unsupported vertical_merge value.");
                }
            }
            if (checked(cursor + row.GridAfter) != table.GridColumns)
                throw Invalid($"Authored document table row {rowIndex} does not cover its declared logical grid.");
            foreach (var (key, group) in active)
                if (!continued.TryGetValue(key, out var carried) || !ReferenceEquals(group, carried)) Finish(group);
            active = continued;
        }
        foreach (var group in active.Values) Finish(group);
    }

    private static W.TableCell BuildCell(
        string value,
        DocumentTableCell? geometry = null,
        DocumentTableFormatting? formatting = null,
        uint gridColumn = 0,
        bool header = false)
    {
        var cell = new W.TableCell();
        var properties = new W.TableCellProperties();
        var width = DocxTableFormatting.BuildCellWidth(formatting, gridColumn, geometry?.ColumnSpan ?? 1);
        if (width is not null) properties.Append(width);
        if (geometry is not null)
        {
            if (geometry.ColumnSpan > 1) properties.Append(new W.GridSpan { Val = checked((int)geometry.ColumnSpan) });
            if (geometry.VerticalMerge == DocumentTableVerticalMerge.Restart)
                properties.Append(new W.VerticalMerge { Val = W.MergedCellValues.Restart });
            else if (geometry.VerticalMerge == DocumentTableVerticalMerge.Continue)
                properties.Append(new W.VerticalMerge { Val = W.MergedCellValues.Continue });
        }
        var shading = DocxTableFormatting.BuildHeaderShading(formatting, header);
        if (shading is not null) properties.Append(shading);
        if (properties.ChildElements.Count > 0) cell.Append(properties);
        var text = new W.Text(value)
        {
            Space = value.Length != value.Trim().Length ? SpaceProcessingModeValues.Preserve : null,
        };
        var run = new W.Run();
        var runProperties = DocxTableFormatting.BuildHeaderRunProperties(formatting, header);
        if (runProperties is not null) run.Append(runProperties);
        run.Append(text);
        cell.Append(new W.Paragraph(run));
        return cell;
    }

    private static void Finish(MergeGroup group)
    {
        if (group.Seen != group.Expected)
            throw Invalid($"Authored merge origin {group.Row},{group.Cell} declares row_span {group.Expected} but spans {group.Seen} rows.");
    }

    private static bool HasSafeContainerTopology(W.Table table)
    {
        if (table.ChildElements.Any(child => child is not W.TableProperties and not W.TableGrid and not W.TableRow)) return false;
        var rows = table.Elements<W.TableRow>().ToArray();
        if (rows.Length == 0) return false;
        foreach (var row in rows)
        {
            if (row.ChildElements.Any(child => child is not W.TableRowProperties and not W.TableCell)) return false;
            var cells = row.Elements<W.TableCell>().ToArray();
            if (cells.Length == 0) return false;
        }
        return true;
    }

    private static string Hash(ReadOnlySpan<byte> bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static CodecException Invalid(string message) => new("invalid_document_table", message);
    private static CodecException Unsupported(string message) => new("unsupported_document_edit", message, "word/document.xml");

    private sealed class MergeGroup(int row, int cell, int expected)
    {
        internal int Row { get; } = row;
        internal int Cell { get; } = cell;
        internal int Expected { get; } = expected;
        internal int Seen { get; set; } = 1;
    }
}
