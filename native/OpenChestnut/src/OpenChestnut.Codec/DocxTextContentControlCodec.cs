using System.Globalization;
using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Owns one deliberately bounded WordprocessingML SDT profile: an inline
// plain-text w:sdt containing exactly one modeled run, with alias, tag, native
// ID, and w:text properties only. Every richer SDT remains opaque/source-bound.
internal static class DocxTextContentControlCodec
{
    private const int MaxTagLength = 64;
    private const int MaxAliasLength = 255;

    internal static void AssignNativeIds(DocumentArtifact document)
    {
        var controls = Controls(document).ToArray();
        var used = controls
            .Where(item => item.Control.HasNativeId && item.Control.NativeId is > 0 and <= int.MaxValue)
            .Select(item => item.Control.NativeId)
            .ToHashSet();
        uint next = 1;
        foreach (var (_, control) in controls)
        {
            if (control.HasNativeId) continue;
            while (used.Contains(next)) next++;
            if (next > int.MaxValue)
                throw new CodecException(
                    "invalid_document_content_control",
                    "Document plain-text content controls exhausted the positive native ID range.");
            control.NativeId = next;
            used.Add(next++);
        }
    }

    internal static void Validate(DocumentArtifact document)
    {
        var modelIds = new HashSet<string>(StringComparer.Ordinal);
        var nativeIds = new HashSet<uint>();
        foreach (var (block, control) in Controls(document))
        {
            ValidateText(control.Id, $"Document block {block.Id} content-control model ID", 1, 255);
            if (!modelIds.Add(control.Id))
                throw new CodecException(
                    "invalid_document_content_control",
                    $"Document content-control model ID {control.Id} is duplicated.");
            ValidateText(control.Tag, $"Document content control {control.Id} tag", 1, MaxTagLength);
            ValidateText(control.Alias, $"Document content control {control.Id} alias", 0, MaxAliasLength);
            if (!control.HasNativeId || control.NativeId is 0 or > int.MaxValue || !nativeIds.Add(control.NativeId))
                throw new CodecException(
                    "invalid_document_content_control",
                    $"Document content control {control.Id} requires a unique native ID from 1 through {int.MaxValue.ToString(CultureInfo.InvariantCulture)}.");
        }
    }

    internal static bool IsSupported(W.SdtRun source)
    {
        var properties = source.SdtProperties;
        var content = source.SdtContentRun;
        if (properties is null || content is null || source.ChildElements.Count != 2) return false;
        if (properties.Elements<W.SdtAlias>().Count() > 1 ||
            properties.Elements<W.Tag>().Count() != 1 ||
            properties.Elements<W.SdtId>().Count() != 1 ||
            properties.Elements<W.SdtContentText>().Count() != 1) return false;
        if (properties.ChildElements.Any(child => child is not W.SdtAlias and not W.Tag and not W.SdtId and not W.SdtContentText)) return false;
        var tag = properties.GetFirstChild<W.Tag>()?.Val?.Value;
        var alias = properties.GetFirstChild<W.SdtAlias>()?.Val?.Value ?? tag;
        var nativeId = properties.GetFirstChild<W.SdtId>()?.Val?.Value;
        var textProperties = properties.GetFirstChild<W.SdtContentText>();
        if (!ValidText(tag, 1, MaxTagLength) || !ValidText(alias, 0, MaxAliasLength) || nativeId is null or <= 0 || textProperties?.MultiLine?.Value == true) return false;
        if (content.ChildElements.Count != 1 || content.FirstChild is not W.Run run) return false;
        return run.ChildElements.All(child => child is W.RunProperties or W.Text) &&
               DocxFormattingCodec.IsSupportedRunProperties(run.RunProperties);
    }

    internal static DocumentRun Read(W.SdtRun source, string modelId)
    {
        if (!IsSupported(source))
            throw new CodecException(
                "unsupported_document_content_control",
                "DOCX inline content control is outside the bounded plain-text SDT profile.",
                "word/document.xml");
        var properties = source.SdtProperties!;
        var result = DocxCodec.ReadRun((W.Run)source.SdtContentRun!.FirstChild!);
        result.TextContentControl = new DocumentTextContentControl
        {
            Id = modelId,
            Tag = properties.GetFirstChild<W.Tag>()!.Val!.Value!,
            Alias = properties.GetFirstChild<W.SdtAlias>()?.Val?.Value ?? properties.GetFirstChild<W.Tag>()!.Val!.Value!,
            NativeId = checked((uint)properties.GetFirstChild<W.SdtId>()!.Val!.Value),
        };
        return result;
    }

    internal static W.SdtRun Build(DocumentRun source)
    {
        var control = source.TextContentControl ?? throw new CodecException(
            "invalid_document_content_control",
            "Document plain-text content-control run has no control metadata.");
        var properties = new W.SdtProperties();
        if (control.Alias.Length > 0) properties.Append(new W.SdtAlias { Val = control.Alias });
        properties.Append(
            new W.Tag { Val = control.Tag },
            new W.SdtId { Val = checked((int)control.NativeId) },
            new W.SdtContentText());
        return new W.SdtRun(properties, new W.SdtContentRun(DocxCodec.BuildRun(source)));
    }

    internal static void AssertTopology(DocumentParagraph requested, DocumentParagraph original, string blockId)
    {
        var requestedControls = Topology(requested);
        var sourceControls = Topology(original);
        if (requestedControls.Count != sourceControls.Count ||
            requestedControls.Where((item, index) => item != sourceControls[index]).Any())
            throw new CodecException(
                "document_content_control_topology_changed",
                $"Imported document paragraph {blockId} plain-text content-control topology is source-bound.",
                "word/document.xml");
    }

    private static List<(int RunIndex, uint NativeId)> Topology(DocumentParagraph paragraph) =>
        paragraph.Runs
            .Select((run, index) => (Run: run, Index: index))
            .Where(item => item.Run.TextContentControl is not null)
            .Select(item => (item.Index, item.Run.TextContentControl.NativeId))
            .ToList();

    private static IEnumerable<(DocumentBlock Block, DocumentTextContentControl Control)> Controls(DocumentArtifact document) =>
        document.Blocks
            .Where(block => block.ContentCase == DocumentBlock.ContentOneofCase.Paragraph)
            .SelectMany(block => block.Paragraph.Runs
                .Where(run => run.TextContentControl is not null)
                .Select(run => (block, run.TextContentControl)));

    private static bool ValidText(string? value, int minimum, int maximum) =>
        value is not null && value.Length >= minimum && value.Length <= maximum &&
        !value.Any(character => char.IsControl(character));

    private static void ValidateText(string? value, string label, int minimum, int maximum)
    {
        if (!ValidText(value, minimum, maximum))
            throw new CodecException(
                "invalid_document_content_control",
                $"{label} must contain {minimum} to {maximum} characters without controls.");
    }
}
