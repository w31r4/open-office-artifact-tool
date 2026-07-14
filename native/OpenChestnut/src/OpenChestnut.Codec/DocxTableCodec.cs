using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Owns cell text only for a fixed, deliberately narrow table topology. Table,
// row, cell, paragraph, and run formatting stay in source XML and are guarded
// by the residual hash. The companion geometry codec exposes gridSpan/vMerge;
// horizontal and restart cells with simple text are editable, while vertical
// continuations and complex cell content remain source-preserved and read-only.
internal static class DocxTableCodec
{
    internal static DocumentTable Read(W.Table table, out bool editable)
    {
        var artifact = DocxTableGeometry.Read(table, out var validGeometry);
        editable = validGeometry && HasSafeContainerTopology(table) &&
                   table.Descendants<W.TableCell>().All(DocxTableGeometry.IsSimpleCell) &&
                   artifact.Rows.SelectMany(row => row.RichCells).Any(cell => cell.Editable);
        return artifact;
    }

    internal static void Apply(W.Table table, DocumentTable requested)
    {
        Validate(requested);
        var source = Read(table, out var editable);
        if (!editable) throw Unsupported("Source-preserving DOCX export cannot edit this table topology.");
        if (!DocxTableGeometry.SameTopology(requested, source))
            throw Unsupported("Source-preserving DOCX table grid, span, and merge topology cannot be changed.");

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
    }

    internal static string ResidualHash(W.Table table)
    {
        var clone = (W.Table)table.CloneNode(true);
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
                if (cell.VerticalMerge != DocumentTableVerticalMerge.Continue && cell.RowSpan == 0)
                    throw Invalid($"Document table origin cell {rowIndex},{cellIndex} must have a positive row_span.");
            }
        }
        if (table.Rows.Any(row => row.RichCells.Count > 0) && table.GridColumns is 0 or > 4_096)
            throw Invalid("Document table grid_columns must be between 1 and 4096 when rich cell geometry is present.");
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
}
