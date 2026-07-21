using System.Security.Cryptography;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Owns the narrowest source-bound paragraph edit: one literal replacement in
// an ordinary w:r/w:t node or an adjacent same-format run span. It
// intentionally does not cross formatting boundaries or enter hyperlinks,
// fields, content controls, revisions, drawings, or other nested run graphs.
internal static class DocxPlainTextPatchCodec
{
    internal static bool IsPatchable(W.Paragraph paragraph) =>
        DocxLiteralTextSpanCodec.IsPatchable(paragraph);

    internal static void Validate(DocumentBlock block)
    {
        if (block.TextPatches.Count == 0) return;
        if (block.ContentCase != DocumentBlock.ContentOneofCase.Paragraph)
            throw Invalid($"Document block {block.Id} text patches are supported only for paragraph content.");
        if (block.Source is null)
            throw Invalid($"Document block {block.Id} text patches require an imported source binding.");
        if (block.TextPatches.Count > 10_000)
            throw Invalid($"Document block {block.Id} exceeds 10,000 source text patches.");
        foreach (var patch in block.TextPatches)
        {
            if (string.IsNullOrEmpty(patch.Search) || patch.Search.Length > 1_000_000 || patch.Replacement.Length > 1_000_000 ||
                !XmlSafe(patch.Search) || !XmlSafe(patch.Replacement))
                throw Invalid($"Document block {block.Id} text patch requires bounded XML-safe search and replacement strings.");
            if (patch.SourceTextSha256.Length != 64 || patch.SourceTextSha256.Any(character => !Uri.IsHexDigit(character)))
                throw Invalid($"Document block {block.Id} text patch requires a SHA-256 source text binding.");
        }
    }

    internal static void Apply(W.Paragraph paragraph, DocumentBlock requested, string sourceText)
    {
        var residual = ResidualHash(paragraph);
        var sourceHash = Hash(Encoding.UTF8.GetBytes(sourceText));
        var expected = sourceText;
        foreach (var patch in requested.TextPatches)
        {
            if (!sourceHash.Equals(patch.SourceTextSha256, StringComparison.OrdinalIgnoreCase))
                throw Unsupported($"Document block {requested.Id} text no longer matches the patch source binding.");
            var expectedOffset = expected.IndexOf(patch.Search, StringComparison.Ordinal);
            if (expectedOffset < 0 || expected.IndexOf(patch.Search, expectedOffset + 1, StringComparison.Ordinal) >= 0)
                throw Unsupported($"Document block {requested.Id} text patch requires exactly one visible match.");
            var resolution = DocxLiteralTextSpanCodec.Resolve(paragraph, expected, patch.Search);
            if (resolution.Status == DocxLiteralTextSpanStatus.TextMismatch)
                throw Unsupported($"Document block {requested.Id} native text no longer matches its semantic source snapshot.");
            if (resolution.Status == DocxLiteralTextSpanStatus.MatchNotUnique)
                throw Unsupported($"Document block {requested.Id} text patch requires exactly one visible match; found {resolution.MatchCount}.");
            if (resolution.Status != DocxLiteralTextSpanStatus.Success || resolution.Span is null)
                throw Unsupported($"Document block {requested.Id} text patch must stay inside one ordinary native text node or adjacent same-format runs; {DocxLiteralTextSpanCodec.FailureDescription(resolution.Status)}.");
            expected = expected.Remove(expectedOffset, patch.Search.Length).Insert(expectedOffset, patch.Replacement);
            DocxLiteralTextSpanCodec.Replace(resolution.Span, patch.Replacement);
        }
        if (!ResidualHash(paragraph).Equals(residual, StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "document_residual_not_preserved",
                $"Document block {requested.Id} text patch changed source-bound paragraph or run topology.",
                "word/document.xml");
        if (!string.Concat(paragraph.Descendants<W.Text>().Select(text => text.Text)).Equals(expected, StringComparison.Ordinal))
            throw new CodecException(
                "document_semantics_not_applied",
                $"Document block {requested.Id} text patches did not produce the requested visible text.",
                "word/document.xml");
    }

    internal static string ResidualHash(W.Paragraph paragraph)
    {
        var clone = (W.Paragraph)paragraph.CloneNode(true);
        foreach (var text in clone.Descendants<W.Text>())
        {
            text.Text = string.Empty;
            text.Space = null;
        }
        return Hash(Encoding.UTF8.GetBytes(clone.OuterXml));
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

    private static string Hash(ReadOnlySpan<byte> bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static CodecException Invalid(string message) => new("invalid_document_text_patch", message, "word/document.xml");
    private static CodecException Unsupported(string message) => new("unsupported_document_edit", message, "word/document.xml");
}
