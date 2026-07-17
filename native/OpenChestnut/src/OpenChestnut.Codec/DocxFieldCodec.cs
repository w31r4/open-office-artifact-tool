using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Owns bounded whole-paragraph fields. Ordinary commands use w:fldSimple. A
// second canonical one-run begin/instrText/separate/result/end topology is
// limited to TOC placeholders; refreshed cross-paragraph TOC result graphs
// remain source-preserved and read-only.
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

    private static readonly Regex CanonicalToc = new(
        "^TOC \\\\o \\\"([1-9])-([1-9])\\\"(?: \\\\h)?(?: \\\\z)?(?: \\\\u)?$",
        RegexOptions.CultureInvariant);

    internal static bool TryRead(W.Paragraph paragraph, out DocumentField field, out bool editable)
    {
        if (TryReadSimple(paragraph, out field, out editable)) return true;
        return TryReadComplex(paragraph, out field, out editable);
    }

    internal static W.Paragraph Build(DocumentBlock block)
    {
        Validate(block.Field);
        var paragraph = new W.Paragraph();
        if (!string.IsNullOrWhiteSpace(block.StyleId))
            paragraph.ParagraphProperties = new W.ParagraphProperties(new W.ParagraphStyleId { Val = block.StyleId });
        if (!block.Field.Complex)
        {
            paragraph.Append(new W.SimpleField(
                new W.Run(Text(block.Field.Display)))
            {
                Instruction = block.Field.Instruction,
            });
            return paragraph;
        }

        paragraph.Append(new W.Run(
            new W.FieldChar { FieldCharType = W.FieldCharValues.Begin },
            FieldCode(block.Field.Instruction),
            new W.FieldChar { FieldCharType = W.FieldCharValues.Separate },
            Text(block.Field.Display),
            new W.FieldChar { FieldCharType = W.FieldCharValues.End }));
        return paragraph;
    }

    internal static void Apply(W.Paragraph paragraph, DocumentField requested)
    {
        Validate(requested);
        if (requested.Complex)
        {
            if (!TryComplexParts(paragraph, out _, out var code, out var result) ||
                !IsEditableComplexInstruction(code.Text ?? string.Empty))
                throw Unsupported("Source-preserving DOCX export cannot edit this complex field instruction or result topology.");
            code.Text = $" {requested.Instruction} ";
            code.Space = SpaceProcessingModeValues.Preserve;
            SetText(result, requested.Display);
            return;
        }

        var source = paragraph.Elements<W.SimpleField>().SingleOrDefault();
        if (source is null || !IsEditableSimple(source) || !IsEditableSimpleInstruction(source.Instruction?.Value ?? string.Empty))
            throw Unsupported("Source-preserving DOCX export cannot edit this simple field instruction or result topology.");
        source.Instruction = requested.Instruction;
        SetText(source.Descendants<W.Text>().Single(), requested.Display);
    }

    internal static string ResidualHash(W.Paragraph paragraph)
    {
        var clone = (W.Paragraph)paragraph.CloneNode(true);
        var field = clone.Elements<W.SimpleField>().SingleOrDefault();
        if (field is not null)
        {
            field.Instruction = string.Empty;
            foreach (var text in field.Descendants<W.Text>()) ClearText(text);
        }
        else if (TryComplexParts(clone, out _, out var code, out var result))
        {
            code.Text = string.Empty;
            code.Space = null;
            ClearText(result);
        }
        return Hash(Encoding.UTF8.GetBytes(clone.OuterXml));
    }

    internal static void Validate(DocumentField? field)
    {
        if (field is null) throw Invalid("Document field payload is missing.");
        ValidateInstruction(field.Instruction);
        if (field.Complex ? !IsCanonicalToc(field.Instruction) : !IsEditableSimpleInstruction(field.Instruction))
            throw Invalid($"Document field command {Command(field.Instruction)} is outside the bounded editable field catalog or topology.");
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

    private static bool TryReadSimple(W.Paragraph paragraph, out DocumentField field, out bool editable)
    {
        field = new DocumentField();
        editable = false;
        if (paragraph.ChildElements.Any(child => child is not W.ParagraphProperties and not W.SimpleField)) return false;
        var fields = paragraph.Elements<W.SimpleField>().ToArray();
        if (fields.Length != 1 || fields[0].Instruction?.Value is not { } instruction) return false;
        var source = fields[0];
        field.Instruction = instruction;
        field.Display = string.Concat(source.Descendants<W.Text>().Select(item => item.Text));
        field.Complex = false;
        editable = IsEditableSimple(source) && IsEditableSimpleInstruction(instruction);
        return true;
    }

    private static bool TryReadComplex(W.Paragraph paragraph, out DocumentField field, out bool editable)
    {
        field = new DocumentField();
        editable = false;
        if (!TryComplexParts(paragraph, out _, out var code, out var result)) return false;
        var instruction = (code.Text ?? string.Empty).Trim();
        field.Instruction = instruction;
        field.Display = result.Text;
        field.Complex = true;
        editable = IsEditableComplexInstruction(instruction);
        return true;
    }

    private static bool TryComplexParts(
        W.Paragraph paragraph,
        out W.Run run,
        out W.FieldCode code,
        out W.Text result)
    {
        run = null!;
        code = null!;
        result = null!;
        if (paragraph.ChildElements.Any(child => child is not W.ParagraphProperties and not W.Run)) return false;
        var runs = paragraph.Elements<W.Run>().ToArray();
        if (runs.Length != 1) return false;
        run = runs[0];
        var content = run.ChildElements.Where(child => child is not W.RunProperties).ToArray();
        if (content.Length != 5 ||
            content[0] is not W.FieldChar begin || begin.FieldCharType?.Value != W.FieldCharValues.Begin ||
            content[1] is not W.FieldCode instruction ||
            content[2] is not W.FieldChar separate || separate.FieldCharType?.Value != W.FieldCharValues.Separate ||
            content[3] is not W.Text text ||
            content[4] is not W.FieldChar end || end.FieldCharType?.Value != W.FieldCharValues.End)
            return false;
        code = instruction;
        result = text;
        return true;
    }

    private static bool IsEditableSimple(W.SimpleField source)
    {
        if (source.ChildElements.Any(child => child is not W.Run)) return false;
        var runs = source.Elements<W.Run>().ToArray();
        if (runs.Length != 1) return false;
        var run = runs[0];
        if (run.ChildElements.Any(child => child is not W.RunProperties and not W.Text)) return false;
        return run.Elements<W.Text>().Count() == 1;
    }

    private static bool IsEditableSimpleInstruction(string value)
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

    private static bool IsEditableComplexInstruction(string value)
    {
        try
        {
            ValidateInstruction(value.Trim());
            return IsCanonicalToc(value.Trim());
        }
        catch (CodecException)
        {
            return false;
        }
    }

    private static bool IsCanonicalToc(string value)
    {
        var match = CanonicalToc.Match(value);
        return match.Success && int.Parse(match.Groups[1].Value) <= int.Parse(match.Groups[2].Value);
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

    private static W.FieldCode FieldCode(string value) => new($" {value} ")
    {
        Space = SpaceProcessingModeValues.Preserve,
    };

    private static W.Text Text(string value) => new(value)
    {
        Space = value.Length != value.Trim().Length ? SpaceProcessingModeValues.Preserve : null,
    };

    private static void SetText(W.Text text, string value)
    {
        text.Text = value;
        text.Space = value.Length != value.Trim().Length ? SpaceProcessingModeValues.Preserve : null;
    }

    private static void ClearText(W.Text text)
    {
        text.Text = string.Empty;
        text.Space = null;
    }

    private static string Hash(ReadOnlySpan<byte> bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static CodecException Invalid(string message) => new("invalid_document_field", message);
    private static CodecException Unsupported(string message) => new("unsupported_document_edit", message, "word/document.xml");
}
