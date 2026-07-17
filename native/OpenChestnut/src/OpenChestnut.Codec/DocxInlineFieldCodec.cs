using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Owns one bounded logical inline field encoded as five consecutive physical
// runs: begin, instrText, separate, cached-result, end. A SEQ field may wrap
// only its cached result in one bookmarkStart/bookmarkEnd pair. The logical
// run's Text is the cached result; its InlineField carries both identities.
internal static class DocxInlineFieldCodec
{
    private static readonly Regex CanonicalInstruction = new(
        "^(?:SEQ [A-Za-z][A-Za-z0-9_]{0,39} \\\\[*] ARABIC|(?:REF|PAGEREF) [A-Za-z][A-Za-z0-9_]{0,39} \\\\h)$",
        RegexOptions.CultureInvariant);

    internal static bool TryRead(
        IReadOnlyList<OpenXmlElement> children,
        int start,
        out DocumentRun run,
        out int consumed)
    {
        run = new DocumentRun();
        consumed = 0;
        if (start < 0 || start + 5 > children.Count ||
            children[start] is not W.Run beginRun ||
            children[start + 1] is not W.Run instructionRun ||
            children[start + 2] is not W.Run separateRun)
            return false;
        var bookmarkName = string.Empty;
        var bookmarkNativeId = string.Empty;
        W.Run resultRun;
        W.Run endRun;
        if (start + 7 <= children.Count &&
            children[start + 3] is W.BookmarkStart bookmarkStart &&
            children[start + 4] is W.Run bookmarkedResultRun &&
            children[start + 5] is W.BookmarkEnd bookmarkEnd &&
            children[start + 6] is W.Run bookmarkedEndRun)
        {
            bookmarkName = bookmarkStart.Name?.Value ?? string.Empty;
            bookmarkNativeId = bookmarkStart.Id?.Value ?? string.Empty;
            if (!DocxBookmarkCodec.ValidName(bookmarkName) ||
                !uint.TryParse(bookmarkNativeId, out _) ||
                bookmarkEnd.Id?.Value != bookmarkNativeId)
                return false;
            resultRun = bookmarkedResultRun;
            endRun = bookmarkedEndRun;
            consumed = 7;
        }
        else
        {
            if (children[start + 3] is not W.Run plainResultRun ||
                children[start + 4] is not W.Run plainEndRun)
                return false;
            resultRun = plainResultRun;
            endRun = plainEndRun;
            consumed = 5;
        }
        if (!SingleFieldChar(beginRun, W.FieldCharValues.Begin) ||
            !SingleFieldCode(instructionRun, out var instruction) ||
            !SingleFieldChar(separateRun, W.FieldCharValues.Separate) ||
            !SingleFieldChar(endRun, W.FieldCharValues.End) ||
            !IsResultRun(resultRun) ||
            !IsCanonical(instruction))
            return false;

        run = DocxCodec.ReadRun(resultRun);
        run.InlineField = new DocumentInlineField
        {
            Instruction = instruction.Trim(),
            BookmarkName = bookmarkName,
            BookmarkNativeId = bookmarkNativeId,
        };
        return true;
    }

    internal static IReadOnlyList<OpenXmlElement> Build(DocumentRun source)
    {
        Validate(source);
        var elements = new List<OpenXmlElement>
        {
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }),
            new W.Run(new W.FieldCode($" {source.InlineField.Instruction} ") { Space = SpaceProcessingModeValues.Preserve }),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate })
        };
        if (source.InlineField.BookmarkName.Length > 0)
            elements.Add(new W.BookmarkStart
            {
                Name = source.InlineField.BookmarkName,
                Id = source.InlineField.BookmarkNativeId,
            });
        elements.Add(DocxCodec.BuildRun(source));
        if (source.InlineField.BookmarkName.Length > 0)
            elements.Add(new W.BookmarkEnd { Id = source.InlineField.BookmarkNativeId });
        elements.Add(new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End }));
        return elements;
    }

    internal static void Validate(DocumentRun source)
    {
        if (source.InlineField is null || !IsCanonical(source.InlineField.Instruction))
            throw new CodecException(
                "invalid_document_inline_field",
                "Document inline field must be canonical SEQ <label> \\* ARABIC, REF <bookmark> \\h, or PAGEREF <bookmark> \\h.");
        if (source.TextContentControl is not null)
            throw new CodecException("invalid_document_inline_field", "Document run cannot combine an inline field and a text content control.");
        if (source.InlineField.BookmarkName.Length > 0)
        {
            if (!source.InlineField.Instruction.TrimStart().StartsWith("SEQ ", StringComparison.Ordinal))
                throw new CodecException("invalid_document_inline_field", "Only a canonical SEQ inline field may carry a cached-result bookmark.");
            if (!DocxBookmarkCodec.ValidName(source.InlineField.BookmarkName) ||
                !uint.TryParse(source.InlineField.BookmarkNativeId, out _))
                throw new CodecException("invalid_document_inline_field", "Document inline field bookmark requires a valid name and planned unsigned native ID.");
        }
        else if (source.InlineField.BookmarkNativeId.Length > 0)
            throw new CodecException("invalid_document_inline_field", "Document inline field bookmark native ID requires a bookmark name.");
        if (source.Text.Length > 1_000_000)
            throw new CodecException("invalid_document_inline_field", "Document inline field cached display exceeds 1,000,000 characters.");
    }

    internal static void AssertTopology(DocumentParagraph requested, DocumentParagraph original, string blockId)
    {
        var requestedFields = Topology(requested);
        var originalFields = Topology(original);
        if (!requestedFields.SequenceEqual(originalFields))
            throw new CodecException(
                "document_inline_field_topology_changed",
                $"Imported document paragraph {blockId} inline-field positions and instructions are source-bound.",
                "word/document.xml");
    }

    internal static bool HasInlineField(IEnumerable<DocumentRun> runs) => runs.Any(run => run.InlineField is not null);

    private static IEnumerable<string> Topology(DocumentParagraph paragraph) =>
        paragraph.Runs.Select((run, index) => (run, index))
            .Where(item => item.run.InlineField is not null)
            .Select(item => $"{item.index}:{item.run.InlineField.Instruction.Trim()}:{item.run.InlineField.BookmarkName}:{item.run.InlineField.BookmarkNativeId}");

    private static bool IsCanonical(string value) =>
        !string.IsNullOrWhiteSpace(value) &&
        value.Length <= 8_192 &&
        !value.Any(char.IsControl) &&
        CanonicalInstruction.IsMatch(value.Trim());

    private static bool SingleFieldChar(W.Run run, W.FieldCharValues expected)
    {
        if (run.RunProperties is not null) return false;
        var content = run.ChildElements.ToArray();
        return content.Length == 1 && content[0] is W.FieldChar field && field.FieldCharType?.Value == expected;
    }

    private static bool SingleFieldCode(W.Run run, out string instruction)
    {
        instruction = string.Empty;
        if (run.RunProperties is not null) return false;
        var content = run.ChildElements.ToArray();
        if (content.Length != 1 || content[0] is not W.FieldCode code) return false;
        instruction = (code.Text ?? string.Empty).Trim();
        return true;
    }

    private static bool IsResultRun(W.Run run) =>
        run.ChildElements.All(child => child is W.RunProperties or W.Text) &&
        run.Elements<W.Text>().Count() == 1 &&
        DocxFormattingCodec.IsSupportedRunProperties(run.RunProperties);
}
