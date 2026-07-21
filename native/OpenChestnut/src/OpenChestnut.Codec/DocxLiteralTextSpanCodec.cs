using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

internal enum DocxLiteralTextSpanStatus
{
    Success,
    TextMismatch,
    MatchNotUnique,
    UnsupportedTopology,
    CrossParagraph,
    EmptyRunGap,
    RunPropertyMismatch,
    MappingFailed,
}

internal sealed record DocxLiteralTextSegment(
    W.Run Run,
    W.Text Text,
    int ChildIndex,
    int Start,
    int Length);

internal sealed record DocxLiteralTextSpan(IReadOnlyList<DocxLiteralTextSegment> Segments);

internal sealed record DocxLiteralTextSpanResolution(
    DocxLiteralTextSpanStatus Status,
    DocxLiteralTextSpan? Span = null,
    int MatchCount = 0);

// Maps one unique visible literal back to its exact ordinary Word runs. Both
// tracked and untracked edits use this representation so adjacency, paragraph
// ownership, empty-run gaps, and formatting boundaries have one definition.
internal static class DocxLiteralTextSpanCodec
{
    internal static bool IsPatchable(W.Paragraph paragraph) =>
        paragraph.Elements<W.Run>().Any(run =>
            TryReadOrdinaryTextRun(run, out var text) && text.Text.Length > 0);

    internal static bool IsPatchable(W.TableCell cell) =>
        cell.Elements<W.Paragraph>().SelectMany(paragraph => paragraph.Elements<W.Run>()).Any(run =>
            TryReadOrdinaryTextRun(run, out var text) && text.Text.Length > 0);

    internal static DocxLiteralTextSpanResolution Resolve(
        OpenXmlElement owner,
        string expectedText,
        string search)
    {
        if (owner is not W.Paragraph and not W.TableCell)
            return new(DocxLiteralTextSpanStatus.UnsupportedTopology);
        if (search.Length == 0)
            return new(DocxLiteralTextSpanStatus.UnsupportedTopology);

        var texts = owner.Descendants<W.Text>().ToArray();
        var actualText = string.Concat(texts.Select(text => text.Text));
        if (!actualText.Equals(expectedText, StringComparison.Ordinal))
            return new(DocxLiteralTextSpanStatus.TextMismatch);

        var matchStart = -1;
        var matchCount = 0;
        for (var index = actualText.IndexOf(search, StringComparison.Ordinal);
             index >= 0;
             index = actualText.IndexOf(search, index + 1, StringComparison.Ordinal))
        {
            if (matchCount == 0) matchStart = index;
            matchCount++;
        }
        if (matchCount != 1)
            return new(DocxLiteralTextSpanStatus.MatchNotUnique, MatchCount: matchCount);

        var matchEnd = checked(matchStart + search.Length);
        var segments = new List<DocxLiteralTextSegment>();
        var textOffset = 0;
        foreach (var text in texts)
        {
            var textStart = textOffset;
            var textEnd = checked(textStart + text.Text.Length);
            var segmentStart = Math.Max(matchStart, textStart);
            var segmentEnd = Math.Min(matchEnd, textEnd);
            if (segmentStart < segmentEnd)
            {
                if (text.Parent is not W.Run run ||
                    !TryReadOrdinaryTextRun(run, out var ordinaryText) ||
                    !ReferenceEquals(ordinaryText, text) ||
                    !IsDirectOwnerRun(owner, run, out var paragraph))
                    return new(DocxLiteralTextSpanStatus.UnsupportedTopology, MatchCount: 1);
                segments.Add(new DocxLiteralTextSegment(
                    run,
                    text,
                    Array.IndexOf(paragraph.ChildElements.ToArray(), run),
                    segmentStart - textStart,
                    segmentEnd - segmentStart));
            }
            textOffset = textEnd;
        }

        if (segments.Count == 0 ||
            string.Concat(segments.Select(segment => segment.Text.Text.Substring(segment.Start, segment.Length))) != search)
            return new(DocxLiteralTextSpanStatus.MappingFailed, MatchCount: 1);

        var parent = segments[0].Run.Parent;
        if (segments.Any(segment => !ReferenceEquals(segment.Run.Parent, parent)))
            return new(DocxLiteralTextSpanStatus.CrossParagraph, MatchCount: 1);
        if (segments.Count > 1 && segments[^1].ChildIndex - segments[0].ChildIndex + 1 != segments.Count)
            return new(DocxLiteralTextSpanStatus.EmptyRunGap, MatchCount: 1);
        if (!SameRunProperties(segments.Select(segment => segment.Run)))
            return new(DocxLiteralTextSpanStatus.RunPropertyMismatch, MatchCount: 1);

        return new(
            DocxLiteralTextSpanStatus.Success,
            new DocxLiteralTextSpan(segments),
            MatchCount: 1);
    }

    internal static void Replace(DocxLiteralTextSpan span, string replacement)
    {
        var first = span.Segments[0];
        var last = span.Segments[^1];
        var prefix = first.Text.Text[..first.Start];
        var suffix = last.Text.Text[(last.Start + last.Length)..];
        if (ReferenceEquals(first.Text, last.Text))
        {
            SetText(first.Text, prefix + replacement + suffix);
            return;
        }

        SetText(first.Text, prefix + replacement);
        for (var index = 1; index < span.Segments.Count - 1; index++)
            SetText(span.Segments[index].Text, string.Empty);
        SetText(last.Text, suffix);
    }

    internal static bool SameRunProperties(IEnumerable<W.Run> runs)
    {
        using var iterator = runs.GetEnumerator();
        if (!iterator.MoveNext()) return false;
        var expected = RunPropertiesKey(iterator.Current);
        while (iterator.MoveNext())
            if (!RunPropertiesKey(iterator.Current).Equals(expected, StringComparison.Ordinal)) return false;
        return true;
    }

    internal static string FailureDescription(DocxLiteralTextSpanStatus status) => status switch
    {
        DocxLiteralTextSpanStatus.TextMismatch => "native text differs from the semantic source snapshot",
        DocxLiteralTextSpanStatus.MatchNotUnique => "the visible literal is not unique",
        DocxLiteralTextSpanStatus.UnsupportedTopology => "the match enters unsupported native text topology",
        DocxLiteralTextSpanStatus.CrossParagraph => "the match crosses a paragraph boundary",
        DocxLiteralTextSpanStatus.EmptyRunGap => "the match crosses an empty-run gap",
        DocxLiteralTextSpanStatus.RunPropertyMismatch => "the matched runs have different formatting",
        DocxLiteralTextSpanStatus.MappingFailed => "the visible literal cannot be mapped back to native text nodes",
        _ => "the literal span is unsupported",
    };

    internal static bool TryReadOrdinaryTextRun(W.Run run, out W.Text text)
    {
        text = null!;
        var properties = run.Elements<W.RunProperties>().ToArray();
        if (properties.Length > 1 ||
            (properties.Length == 1 && !ReferenceEquals(run.FirstChild, properties[0])) ||
            run.ChildElements.Any(child => child is not W.RunProperties and not W.Text))
            return false;
        var texts = run.Elements<W.Text>().ToArray();
        if (texts.Length != 1) return false;
        text = texts[0];
        return true;
    }

    private static bool IsDirectOwnerRun(
        OpenXmlElement owner,
        W.Run run,
        out W.Paragraph paragraph)
    {
        paragraph = run.Parent as W.Paragraph ?? null!;
        if (paragraph is null) return false;
        return owner switch
        {
            W.Paragraph ownerParagraph => ReferenceEquals(paragraph, ownerParagraph),
            W.TableCell cell => ReferenceEquals(paragraph.Parent, cell),
            _ => false,
        };
    }

    private static string RunPropertiesKey(W.Run run) => run.RunProperties?.OuterXml ?? string.Empty;

    private static void SetText(W.Text text, string value)
    {
        text.Text = value;
        text.Space = value.Length != value.Trim().Length ? SpaceProcessingModeValues.Preserve : null;
    }
}
