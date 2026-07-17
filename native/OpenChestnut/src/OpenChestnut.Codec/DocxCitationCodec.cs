using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Owns one reversible whole-paragraph w:fldSimple CITATION profile. The tag is
// fixed for an imported field; result text may change while run formatting,
// bookmarks, dirty state, and every other native detail stay hash-bound.
internal static partial class DocxCitationCodec
{
    [GeneratedRegex("^\\s*CITATION\\s+([A-Za-z0-9_.:-]{1,255})\\s*$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex InstructionPattern();

    internal static bool TryRead(
        W.Paragraph paragraph,
        IReadOnlySet<string> sourceTags,
        out DocumentCitation citation,
        out bool editable)
    {
        citation = new DocumentCitation();
        editable = false;
        if (paragraph.ChildElements.Any(child => child is not W.ParagraphProperties and
                                                 not W.BookmarkStart and
                                                 not W.BookmarkEnd and
                                                 not W.SimpleField)) return false;
        var fields = paragraph.Elements<W.SimpleField>().ToArray();
        if (fields.Length != 1 || fields[0].Instruction?.Value is not { } instruction) return false;
        var match = InstructionPattern().Match(instruction);
        if (!match.Success || !sourceTags.Contains(match.Groups[1].Value)) return false;
        citation.Tag = match.Groups[1].Value;
        citation.Display = string.Concat(fields[0].Descendants<W.Text>().Select(item => item.Text));
        editable = IsEditable(fields[0]);
        return true;
    }

    internal static W.Paragraph Build(DocumentBlock block)
    {
        Validate(block.Citation);
        var paragraph = new W.Paragraph();
        if (!string.IsNullOrWhiteSpace(block.StyleId))
            paragraph.ParagraphProperties = new W.ParagraphProperties(new W.ParagraphStyleId { Val = block.StyleId });
        paragraph.Append(new W.SimpleField(
            new W.Run(Text(block.Citation.Display)))
        {
            Instruction = Instruction(block.Citation.Tag),
            Dirty = true,
        });
        return paragraph;
    }

    internal static void Apply(W.Paragraph paragraph, DocumentCitation requested, DocumentCitation original)
    {
        Validate(requested);
        if (!requested.Tag.Equals(original.Tag, StringComparison.Ordinal))
            throw Unsupported("Imported DOCX citation tags are source-bound; update the source catalog and citation graph in a new document instead.");
        var source = paragraph.Elements<W.SimpleField>().SingleOrDefault();
        if (source is null || !IsEditable(source) ||
            !InstructionPattern().Match(source.Instruction?.Value ?? string.Empty).Success)
            throw Unsupported("Source-preserving DOCX export cannot edit this citation field topology.");
        source.Instruction = Instruction(requested.Tag);
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

    internal static void Validate(DocumentCitation? citation)
    {
        if (citation is null) throw Invalid("Document citation payload is missing.");
        if (!ValidTag(citation.Tag))
            throw Invalid("Document citation tag must contain 1 through 255 ASCII letters, digits, periods, underscores, colons, or hyphens.");
        if (citation.Display.Length > 1_000_000)
            throw Invalid("Document citation display text exceeds 1,000,000 characters.");
    }

    internal static bool ValidTag(string value) =>
        !string.IsNullOrEmpty(value) && InstructionPattern().IsMatch($"CITATION {value}");

    private static bool IsEditable(W.SimpleField source)
    {
        if (source.ChildElements.Any(child => child is not W.Run)) return false;
        var runs = source.Elements<W.Run>().ToArray();
        if (runs.Length != 1) return false;
        var run = runs[0];
        if (run.ChildElements.Any(child => child is not W.RunProperties and not W.Text)) return false;
        return run.Elements<W.Text>().Count() == 1;
    }

    private static string Instruction(string tag) => $" CITATION {tag} ";

    private static W.Text Text(string value) => new(value)
    {
        Space = value.Length != value.Trim().Length ? SpaceProcessingModeValues.Preserve : null,
    };

    private static string Hash(ReadOnlySpan<byte> bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static CodecException Invalid(string message) => new("invalid_document_citation", message, "word/document.xml");
    private static CodecException Unsupported(string message) => new("unsupported_document_edit", message, "word/document.xml");
}
