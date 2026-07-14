using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenOffice.OpenXmlCodec;

// Owns a bounded whole-paragraph w:fldSimple. Complex begin/separate/end
// fields and instructions that can reference or fetch external content remain
// source-preserved and read-only.
internal static class DocxFieldCodec
{
    private static readonly HashSet<string> EditableCommands = new(StringComparer.OrdinalIgnoreCase)
    {
        "PAGE",
        "NUMPAGES",
        "SECTION",
        "SECTIONPAGES",
        "DATE",
        "TIME",
        "CREATEDATE",
        "SAVEDATE",
        "PRINTDATE",
        "AUTHOR",
        "TITLE",
        "SUBJECT",
        "COMMENTS",
        "FILENAME",
        "FILESIZE",
        "NUMWORDS",
        "NUMCHARS",
    };

    internal static bool TryRead(W.Paragraph paragraph, out DocumentField field, out bool editable)
    {
        field = new DocumentField();
        editable = false;
        if (paragraph.ChildElements.Any(child => child is not W.ParagraphProperties and not W.SimpleField)) return false;
        var fields = paragraph.Elements<W.SimpleField>().ToArray();
        if (fields.Length != 1 || fields[0].Instruction?.Value is not { } instruction) return false;
        var source = fields[0];
        field.Instruction = instruction;
        field.Display = string.Concat(source.Descendants<W.Text>().Select(item => item.Text));
        editable = IsEditable(source) && IsEditableInstruction(instruction);
        return true;
    }

    internal static W.Paragraph Build(DocumentBlock block)
    {
        Validate(block.Field);
        var paragraph = new W.Paragraph();
        if (!string.IsNullOrWhiteSpace(block.StyleId))
            paragraph.ParagraphProperties = new W.ParagraphProperties(new W.ParagraphStyleId { Val = block.StyleId });
        paragraph.Append(new W.SimpleField(
            new W.Run(Text(block.Field.Display)))
        {
            Instruction = block.Field.Instruction,
        });
        return paragraph;
    }

    internal static void Apply(W.Paragraph paragraph, DocumentField requested)
    {
        Validate(requested);
        var source = paragraph.Elements<W.SimpleField>().SingleOrDefault();
        if (source is null || !IsEditable(source) || !IsEditableInstruction(source.Instruction?.Value ?? string.Empty))
            throw Unsupported("Source-preserving DOCX export cannot edit this field instruction or result topology.");
        source.Instruction = requested.Instruction;
        var text = source.Descendants<W.Text>().Single();
        text.Text = requested.Display;
        text.Space = requested.Display.Length != requested.Display.Trim().Length ? SpaceProcessingModeValues.Preserve : null;
    }

    internal static string ResidualHash(W.Paragraph paragraph)
    {
        var clone = (W.Paragraph)paragraph.CloneNode(true);
        var field = clone.Elements<W.SimpleField>().SingleOrDefault();
        if (field is not null)
        {
            field.Instruction = string.Empty;
            foreach (var text in field.Descendants<W.Text>())
            {
                text.Text = string.Empty;
                text.Space = null;
            }
        }
        return Hash(Encoding.UTF8.GetBytes(clone.OuterXml));
    }

    internal static void Validate(DocumentField? field)
    {
        if (field is null) throw Invalid("Document field payload is missing.");
        ValidateInstruction(field.Instruction);
        if (!IsEditableInstruction(field.Instruction))
            throw Invalid($"Document field command {Command(field.Instruction)} is outside the bounded editable field catalog.");
        if (field.Display.Length > 1_000_000)
            throw Invalid("Document field display text exceeds 1,000,000 characters.");
    }

    internal static void ValidatePreserved(DocumentField? field)
    {
        if (field is null) throw Invalid("Document field payload is missing.");
        ValidateInstruction(field.Instruction);
        if (field.Display.Length > 1_000_000)
            throw Invalid("Document field display text exceeds 1,000,000 characters.");
    }

    private static bool IsEditable(W.SimpleField source)
    {
        if (source.ChildElements.Any(child => child is not W.Run)) return false;
        var runs = source.Elements<W.Run>().ToArray();
        if (runs.Length != 1) return false;
        var run = runs[0];
        if (run.ChildElements.Any(child => child is not W.RunProperties and not W.Text)) return false;
        return run.Elements<W.Text>().Count() == 1;
    }

    private static bool IsEditableInstruction(string value)
    {
        try
        {
            ValidateInstruction(value);
            return EditableCommands.Contains(Command(value));
        }
        catch (CodecException)
        {
            return false;
        }
    }

    private static string Command(string value)
    {
        var trimmed = value.TrimStart();
        var length = 0;
        while (length < trimmed.Length && char.IsAsciiLetter(trimmed[length])) length++;
        return length == 0 ? "(missing)" : trimmed[..length].ToUpperInvariant();
    }

    private static void ValidateInstruction(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 8_192 || value.Any(character => char.IsControl(character)))
            throw Invalid("Document field instruction must contain 1 through 8192 characters without controls.");
        if (Command(value) == "(missing)")
            throw Invalid("Document field instruction must start with an ASCII command name.");
    }

    private static W.Text Text(string value) => new(value)
    {
        Space = value.Length != value.Trim().Length ? SpaceProcessingModeValues.Preserve : null,
    };

    private static string Hash(ReadOnlySpan<byte> bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static CodecException Invalid(string message) => new("invalid_document_field", message);
    private static CodecException Unsupported(string message) => new("unsupported_document_edit", message, "word/document.xml");
}
