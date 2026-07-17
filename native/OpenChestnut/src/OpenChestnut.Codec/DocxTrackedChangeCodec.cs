using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Owns one whole-paragraph w:ins or w:del containing exactly one text run.
// Mixed normal/revision runs, nested revisions, moves, and property changes
// remain source-preserved and read-only.
internal static class DocxTrackedChangeCodec
{
    internal static bool TryRead(
        W.Paragraph paragraph,
        out DocumentChange change,
        out string nativeId,
        out bool editable)
    {
        change = new DocumentChange();
        nativeId = string.Empty;
        editable = false;
        if (paragraph.ChildElements.Any(child =>
                child is not W.ParagraphProperties and not W.InsertedRun and not W.DeletedRun))
            return false;

        var revisions = paragraph.ChildElements
            .Where(child => child is W.InsertedRun or W.DeletedRun)
            .ToArray();
        if (revisions.Length != 1) return false;

        switch (revisions[0])
        {
            case W.InsertedRun insertion:
                change.Type = DocumentChangeType.Insert;
                change.Author = insertion.Author?.Value ?? string.Empty;
                if (insertion.Date?.Value is { } insertionDate)
                    change.Date = insertionDate.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture);
                nativeId = insertion.Id?.Value ?? string.Empty;
                editable = TryReadRun(insertion, deleted: false, out var insertedText);
                change.Text = insertedText;
                break;
            case W.DeletedRun deletion:
                change.Type = DocumentChangeType.Delete;
                change.Author = deletion.Author?.Value ?? string.Empty;
                if (deletion.Date?.Value is { } deletionDate)
                    change.Date = deletionDate.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture);
                nativeId = deletion.Id?.Value ?? string.Empty;
                editable = TryReadRun(deletion, deleted: true, out var deletedText);
                change.Text = deletedText;
                break;
        }

        editable = editable && nativeId.Length > 0 && change.Author.Length > 0;
        return true;
    }

    internal static W.Paragraph Build(DocumentBlock block, string revisionId)
    {
        Validate(block.Change);
        if (string.IsNullOrWhiteSpace(revisionId))
            throw Invalid("Document tracked change requires a non-empty native revision ID.");

        var paragraph = new W.Paragraph();
        if (!string.IsNullOrWhiteSpace(block.StyleId))
            paragraph.ParagraphProperties = new W.ParagraphProperties(new W.ParagraphStyleId { Val = block.StyleId });

        var run = new W.Run(block.Change.Type == DocumentChangeType.Delete
            ? DeletedText(block.Change.Text)
            : Text(block.Change.Text));
        if (block.Change.Type == DocumentChangeType.Delete)
        {
            var deletion = new W.DeletedRun(run)
            {
                Id = revisionId,
                Author = block.Change.Author,
            };
            deletion.Date = Date(block.Change.Date);
            paragraph.Append(deletion);
        }
        else
        {
            var insertion = new W.InsertedRun(run)
            {
                Id = revisionId,
                Author = block.Change.Author,
            };
            insertion.Date = Date(block.Change.Date);
            paragraph.Append(insertion);
        }
        return paragraph;
    }

    internal static void Apply(W.Paragraph paragraph, DocumentChange requested)
    {
        Validate(requested);
        if (!TryRead(paragraph, out _, out _, out var editable) || !editable)
            throw Unsupported("Source-preserving DOCX export cannot edit this tracked-change topology.");

        switch (paragraph.ChildElements.Single(child => child is W.InsertedRun or W.DeletedRun))
        {
            case W.InsertedRun insertion when requested.Type == DocumentChangeType.Insert:
                insertion.Author = requested.Author;
                insertion.Date = Date(requested.Date);
                SetRunText(insertion, requested.Text, deleted: false);
                break;
            case W.DeletedRun deletion when requested.Type == DocumentChangeType.Delete:
                deletion.Author = requested.Author;
                deletion.Date = Date(requested.Date);
                SetRunText(deletion, requested.Text, deleted: true);
                break;
            default:
                throw Unsupported("Source-preserving DOCX export cannot change tracked insertion/deletion kind.");
        }
    }

    internal static string ResidualHash(W.Paragraph paragraph)
    {
        var clone = (W.Paragraph)paragraph.CloneNode(true);
        switch (clone.ChildElements.SingleOrDefault(child => child is W.InsertedRun or W.DeletedRun))
        {
            case W.InsertedRun insertion:
                insertion.Author = string.Empty;
                insertion.Date = null;
                SetRunText(insertion, string.Empty, deleted: false);
                break;
            case W.DeletedRun deletion:
                deletion.Author = string.Empty;
                deletion.Date = null;
                SetRunText(deletion, string.Empty, deleted: true);
                break;
        }
        return Hash(Encoding.UTF8.GetBytes(clone.OuterXml));
    }

    internal static void Validate(DocumentChange? change)
    {
        ValidatePreserved(change);
        if (string.IsNullOrWhiteSpace(change!.Author))
            throw Invalid("Document tracked-change author is required.");
    }

    internal static void ValidatePreserved(DocumentChange? change)
    {
        if (change is null) throw Invalid("Document tracked-change payload is missing.");
        if (change.Type is not DocumentChangeType.Insert and not DocumentChangeType.Delete)
            throw Invalid("Document tracked-change type must be insert or delete.");
        if (change.Text.Length > 1_000_000)
            throw Invalid("Document tracked-change text exceeds 1,000,000 characters.");
        if (change.Author.Length > 255 || change.Author.Any(char.IsControl))
            throw Invalid("Document tracked-change author must contain at most 255 characters without controls.");
        if (change.HasDate && !DateTimeOffset.TryParse(
                change.Date,
                CultureInfo.InvariantCulture,
                DateTimeStyles.RoundtripKind,
                out _))
            throw Invalid("Document tracked-change date must be an ISO 8601 timestamp.");
    }

    private static bool TryReadRun(OpenXmlCompositeElement revision, bool deleted, out string text)
    {
        text = deleted
            ? string.Concat(revision.Descendants<W.DeletedText>().Select(value => value.Text))
            : string.Concat(revision.Descendants<W.Text>().Select(value => value.Text));
        if (revision.ChildElements.Any(child => child is not W.Run)) return false;
        var runs = revision.Elements<W.Run>().ToArray();
        if (runs.Length != 1) return false;
        var run = runs[0];
        if (run.ChildElements.Any(child => child is not W.RunProperties &&
                (deleted ? child is not W.DeletedText : child is not W.Text)))
            return false;
        if (deleted)
        {
            var values = run.Elements<W.DeletedText>().ToArray();
            if (values.Length != 1) return false;
        }
        else
        {
            var values = run.Elements<W.Text>().ToArray();
            if (values.Length != 1) return false;
        }
        return true;
    }

    private static void SetRunText(OpenXmlCompositeElement revision, string value, bool deleted)
    {
        var run = revision.Elements<W.Run>().Single();
        if (deleted)
        {
            var text = run.Elements<W.DeletedText>().Single();
            text.Text = value;
            text.Space = value.Length != value.Trim().Length ? SpaceProcessingModeValues.Preserve : null;
        }
        else
        {
            var text = run.Elements<W.Text>().Single();
            text.Text = value;
            text.Space = value.Length != value.Trim().Length ? SpaceProcessingModeValues.Preserve : null;
        }
    }

    private static DateTime? Date(string value) => string.IsNullOrEmpty(value)
        ? null
        : DateTimeOffset.Parse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind).UtcDateTime;

    private static W.Text Text(string value) => new(value)
    {
        Space = value.Length != value.Trim().Length ? SpaceProcessingModeValues.Preserve : null,
    };

    private static W.DeletedText DeletedText(string value) => new(value)
    {
        Space = value.Length != value.Trim().Length ? SpaceProcessingModeValues.Preserve : null,
    };

    private static string Hash(ReadOnlySpan<byte> bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static CodecException Invalid(string message) => new("invalid_document_change", message);
    private static CodecException Unsupported(string message) => new("unsupported_document_edit", message, "word/document.xml");
}
