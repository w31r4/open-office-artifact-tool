using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Owns cell text only for a fixed, deliberately narrow table topology. Table,
// row, cell, paragraph, and run formatting stay in source XML and are guarded
// by the residual hash. Merged, nested, multi-paragraph, or multi-run cells are
// preserved but remain read-only.
internal static class DocxTableCodec
{
    internal static DocumentTable Read(W.Table table, out bool editable)
    {
        var artifact = new DocumentTable();
        foreach (var row in table.Elements<W.TableRow>())
        {
            var targetRow = new DocumentTableRow();
            foreach (var cell in row.Elements<W.TableCell>())
                targetRow.Cells.Add(string.Concat(cell.Descendants<W.Text>().Select(text => text.Text)));
            artifact.Rows.Add(targetRow);
        }
        editable = IsEditable(table);
        return artifact;
    }

    internal static void Apply(W.Table table, DocumentTable requested)
    {
        Validate(requested);
        if (!IsEditable(table))
            throw Unsupported("Source-preserving DOCX export cannot edit this table topology.");

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
        }
    }

    private static bool IsEditable(W.Table table)
    {
        if (table.ChildElements.Any(child => child is not W.TableProperties and not W.TableGrid and not W.TableRow)) return false;
        var rows = table.Elements<W.TableRow>().ToArray();
        if (rows.Length == 0) return false;
        foreach (var row in rows)
        {
            if (row.ChildElements.Any(child => child is not W.TableRowProperties and not W.TableCell)) return false;
            var cells = row.Elements<W.TableCell>().ToArray();
            if (cells.Length == 0) return false;
            foreach (var cell in cells)
            {
                if (cell.ChildElements.Any(child => child is not W.TableCellProperties and not W.Paragraph)) return false;
                var paragraphs = cell.Elements<W.Paragraph>().ToArray();
                if (paragraphs.Length != 1) return false;
                var paragraph = paragraphs[0];
                if (paragraph.ChildElements.Any(child => child is not W.ParagraphProperties and not W.Run)) return false;
                var runs = paragraph.Elements<W.Run>().ToArray();
                if (runs.Length != 1) return false;
                var run = runs[0];
                if (run.ChildElements.Any(child => child is not W.RunProperties and not W.Text)) return false;
                if (run.Elements<W.Text>().Count() != 1) return false;
            }
        }
        return true;
    }

    private static string Hash(ReadOnlySpan<byte> bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static CodecException Invalid(string message) => new("invalid_document_table", message);
    private static CodecException Unsupported(string message) => new("unsupported_document_edit", message, "word/document.xml");
}
