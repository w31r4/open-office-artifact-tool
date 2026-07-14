using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenOffice.OpenXmlCodec;

// Models a deliberately bounded Word hyperlink paragraph while preserving its
// paragraph/run formatting in the source element. Only the hyperlink target,
// visible text, tooltip, and history flag are edited in place.
internal static class DocxHyperlinkCodec
{
    internal static bool TryRead(
        W.Paragraph paragraph,
        DocxPartContext context,
        out DocumentHyperlink hyperlink,
        out bool editable)
    {
        hyperlink = new DocumentHyperlink();
        editable = false;
        if (paragraph.ChildElements.Any(child => child is not W.ParagraphProperties and not W.Hyperlink)) return false;
        var source = paragraph.Elements<W.Hyperlink>().SingleOrDefault();
        if (source is null || !TryRead(source, context, out hyperlink)) return false;
        editable = IsEditable(source);
        return true;
    }

    internal static W.Paragraph Build(DocumentBlock block, DocxPartContext context)
    {
        Validate(block.Hyperlink);
        var paragraph = new W.Paragraph();
        if (!string.IsNullOrWhiteSpace(block.StyleId))
            paragraph.ParagraphProperties = new W.ParagraphProperties(new W.ParagraphStyleId { Val = block.StyleId });
        var hyperlink = new W.Hyperlink();
        ApplyTarget(hyperlink, block.Hyperlink, context);
        ApplyMetadata(hyperlink, block.Hyperlink);
        hyperlink.Append(new W.Run(
            new W.RunProperties(
                new W.Color { Val = "0000FF" },
                new W.Underline { Val = W.UnderlineValues.Single }),
            Text(block.Hyperlink.Text)));
        paragraph.Append(hyperlink);
        return paragraph;
    }

    internal static void Apply(
        W.Paragraph paragraph,
        DocumentHyperlink requested,
        DocumentHyperlink original,
        DocxPartContext context)
    {
        Validate(requested);
        var source = paragraph.Elements<W.Hyperlink>().SingleOrDefault();
        if (source is null || !IsEditable(source) || !TryRead(source, context, out var current))
            throw Unsupported("Source-preserving DOCX export cannot edit this hyperlink paragraph topology.");
        if (!requested.RelationshipId.Equals(original.RelationshipId, StringComparison.Ordinal))
            throw Unsupported("DOCX hyperlink relationship IDs are source locators and cannot be edited directly.");

        ApplyTarget(source, requested, context);
        ApplyMetadata(source, requested);
        var text = source.Descendants<W.Text>().Single();
        text.Text = requested.Text;
        text.Space = requested.Text.Length != requested.Text.Trim().Length ? SpaceProcessingModeValues.Preserve : null;
        context.RemoveIfUnreferenced(current.RelationshipId);
    }

    internal static string ResidualHash(W.Paragraph paragraph)
    {
        var clone = (W.Paragraph)paragraph.CloneNode(true);
        var hyperlink = clone.Elements<W.Hyperlink>().SingleOrDefault();
        if (hyperlink is not null)
        {
            hyperlink.Id = null;
            hyperlink.Anchor = null;
            hyperlink.Tooltip = null;
            hyperlink.History = null;
            foreach (var text in hyperlink.Descendants<W.Text>())
            {
                text.Text = string.Empty;
                text.Space = null;
            }
        }
        return Hash(Encoding.UTF8.GetBytes(clone.OuterXml));
    }

    internal static void Validate(DocumentHyperlink? hyperlink)
    {
        if (hyperlink is null) throw Invalid("Document hyperlink payload is missing.");
        switch (hyperlink.TargetCase)
        {
            case DocumentHyperlink.TargetOneofCase.ExternalUri:
                ValidateExternalUri(hyperlink.ExternalUri);
                break;
            case DocumentHyperlink.TargetOneofCase.InternalAnchor:
                ValidateAnchor(hyperlink.InternalAnchor);
                break;
            default:
                throw Invalid("Document hyperlink requires exactly one external URI or internal anchor target.");
        }
        if (hyperlink.Text.Length > 1_000_000)
            throw Invalid("Document hyperlink text exceeds 1,000,000 characters.");
        if (hyperlink.HasTooltip && hyperlink.Tooltip.Length > 260)
            throw Invalid("Document hyperlink tooltip exceeds 260 characters.");
    }

    private static bool TryRead(W.Hyperlink source, DocxPartContext context, out DocumentHyperlink hyperlink)
    {
        hyperlink = new DocumentHyperlink { Text = string.Concat(source.Descendants<W.Text>().Select(item => item.Text)) };
        var relationshipId = source.Id?.Value ?? string.Empty;
        var anchor = source.Anchor?.Value ?? string.Empty;
        if (relationshipId.Length > 0 && anchor.Length == 0)
        {
            if (!context.TryReadExternal(relationshipId, out var uri)) return false;
            hyperlink.ExternalUri = uri;
            hyperlink.RelationshipId = relationshipId;
        }
        else if (relationshipId.Length == 0 && anchor.Length > 0)
        {
            if (!context.HasBookmark(anchor)) return false;
            hyperlink.InternalAnchor = anchor;
        }
        else
        {
            return false;
        }
        if (source.Tooltip?.Value is { } tooltip) hyperlink.Tooltip = tooltip;
        if (source.History?.Value is { } history) hyperlink.History = history;
        try
        {
            Validate(hyperlink);
            return true;
        }
        catch (CodecException)
        {
            hyperlink = new DocumentHyperlink();
            return false;
        }
    }

    private static bool IsEditable(W.Hyperlink source)
    {
        if (source.ChildElements.Any(child => child is not W.Run)) return false;
        var runs = source.Elements<W.Run>().ToArray();
        if (runs.Length != 1) return false;
        var run = runs[0];
        if (run.ChildElements.Any(child => child is not W.RunProperties and not W.Text)) return false;
        return run.Elements<W.Text>().Count() == 1;
    }

    private static void ApplyTarget(W.Hyperlink source, DocumentHyperlink requested, DocxPartContext context)
    {
        var oldRelationshipId = source.Id?.Value ?? string.Empty;
        switch (requested.TargetCase)
        {
            case DocumentHyperlink.TargetOneofCase.ExternalUri:
                source.Id = context.EnsureExternal(requested.ExternalUri, oldRelationshipId);
                source.Anchor = null;
                break;
            case DocumentHyperlink.TargetOneofCase.InternalAnchor:
                if (!context.HasBookmark(requested.InternalAnchor))
                    throw Invalid($"Document hyperlink anchor {requested.InternalAnchor} does not resolve to a bookmark in word/document.xml.");
                source.Id = null;
                source.Anchor = requested.InternalAnchor;
                break;
            default:
                throw Invalid("Document hyperlink requires exactly one external URI or internal anchor target.");
        }
        if (source.Id?.Value != oldRelationshipId) context.RemoveIfUnreferenced(oldRelationshipId);
    }

    private static void ApplyMetadata(W.Hyperlink source, DocumentHyperlink requested)
    {
        source.Tooltip = requested.HasTooltip ? requested.Tooltip : null;
        source.History = requested.HasHistory ? requested.History : null;
    }

    private static W.Text Text(string value) => new(value)
    {
        Space = value.Length != value.Trim().Length ? SpaceProcessingModeValues.Preserve : null,
    };

    private static void ValidateExternalUri(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 4_096 || value.Any(char.IsControl) ||
            !Uri.TryCreate(value, UriKind.Absolute, out var uri) ||
            uri.Scheme is not ("http" or "https"))
            throw Invalid("Document hyperlink URI must be an absolute http(s) URI of at most 4096 characters without controls.");
    }

    private static void ValidateAnchor(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 255 || value.Any(char.IsControl))
            throw Invalid("Document hyperlink anchor must contain 1 through 255 characters without controls.");
    }

    private static string Hash(ReadOnlySpan<byte> bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static CodecException Invalid(string message) => new("invalid_document_hyperlink", message);
    private static CodecException Unsupported(string message) => new("unsupported_document_edit", message, "word/document.xml");
}
