using System.Security.Cryptography;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Owns the narrowest source-bound paragraph edit: one literal replacement in
// one ordinary w:r/w:t node. It intentionally does not interpret arbitrary
// InnerText, cross formatting boundaries, or enter hyperlinks, fields,
// content controls, revisions, drawings, or other nested run graphs.
internal static class DocxPlainTextPatchCodec
{
    internal static bool IsPatchable(W.Paragraph paragraph) => PatchableTexts(paragraph).Any();

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
            var matches = new List<(W.Text Text, int Offset)>();
            foreach (var text in PatchableTexts(paragraph))
            {
                var candidateOffset = text.Text.IndexOf(patch.Search, StringComparison.Ordinal);
                if (candidateOffset < 0) continue;
                if (text.Text.IndexOf(patch.Search, candidateOffset + 1, StringComparison.Ordinal) >= 0)
                    throw Unsupported($"Document block {requested.Id} text patch is ambiguous within one native text node.");
                matches.Add((text, candidateOffset));
            }
            if (matches.Count != 1)
                throw Unsupported($"Document block {requested.Id} text patch must match exactly one plain native text node; found {matches.Count}.");
            var expectedOffset = expected.IndexOf(patch.Search, StringComparison.Ordinal);
            if (expectedOffset < 0 || expected.IndexOf(patch.Search, expectedOffset + 1, StringComparison.Ordinal) >= 0)
                throw Unsupported($"Document block {requested.Id} text patch requires exactly one visible match.");
            expected = expected.Remove(expectedOffset, patch.Search.Length).Insert(expectedOffset, patch.Replacement);
            var (target, offset) = matches[0];
            var value = target.Text.Remove(offset, patch.Search.Length).Insert(offset, patch.Replacement);
            target.Text = value;
            target.Space = value.Length != value.Trim().Length ? SpaceProcessingModeValues.Preserve : null;
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

    private static IEnumerable<W.Text> PatchableTexts(W.Paragraph paragraph) =>
        paragraph.Descendants<W.Text>().Where(text =>
            text.Parent is W.Run run &&
            ReferenceEquals(run.Parent, paragraph) &&
            run.Elements<W.Text>().Count() == 1 &&
            run.ChildElements.All(child => child is W.RunProperties or W.Text));

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
