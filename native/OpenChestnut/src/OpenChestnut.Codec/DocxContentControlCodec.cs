using System.Globalization;
using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Owns three deliberately bounded WordprocessingML SDT profiles: inline
// plain-text, canonical drop-down, and Word 2010+ checkbox controls. Every
// profile contains exactly one modeled run with alias, tag, and native ID.
// Checkbox symbols and drop-down visible text are codec-owned. Every richer
// SDT remains opaque/source-bound.
internal static class DocxContentControlCodec
{
    private const int MaxTagLength = 64;
    private const int MaxAliasLength = 255;
    private const int MaxChoiceCount = 256;
    private const int MaxChoiceTextLength = 255;
    private const string CheckboxFont = "MS Gothic";
    internal const string CheckboxNamespace = "http://schemas.microsoft.com/office/word/2010/wordml";
    private const string CheckedSymbolCode = "2612";
    private const string UncheckedSymbolCode = "2610";
    private const string CheckedGlyph = "☒";
    private const string UncheckedGlyph = "☐";

    internal static void AssignNativeIds(DocumentArtifact document)
    {
        var controls = Controls(document).ToArray();
        var used = controls
            .Where(item => item.Control.HasNativeId && item.Control.NativeId is > 0 and <= int.MaxValue)
            .Select(item => item.Control.NativeId)
            .ToHashSet();
        uint next = 1;
        foreach (var (_, _, control) in controls)
        {
            if (control.HasNativeId) continue;
            while (used.Contains(next)) next++;
            if (next > int.MaxValue)
                throw new CodecException(
                    "invalid_document_content_control",
                    "Document content controls exhausted the positive native ID range.");
            control.NativeId = next;
            used.Add(next++);
        }
    }

    internal static bool UsesCheckboxes(DocumentArtifact document) =>
        Controls(document).Any(item => NormalizedType(item.Control) == DocumentContentControlType.Checkbox);

    internal static void Validate(DocumentArtifact document)
    {
        var modelIds = new HashSet<string>(StringComparer.Ordinal);
        var nativeIds = new HashSet<uint>();
        foreach (var (block, run, control) in Controls(document))
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
            var controlType = NormalizedType(control);
            if (controlType == DocumentContentControlType.Checkbox &&
                !run.Text.Equals(control.Checked ? CheckedGlyph : UncheckedGlyph, StringComparison.Ordinal))
                throw new CodecException(
                    "invalid_document_content_control",
                    $"Document checkbox content control {control.Id} visible glyph does not match its checked state.");
            if (controlType == DocumentContentControlType.DropDown)
            {
                var selected = ValidateDropdown(control);
                if (!run.Text.Equals(selected.DisplayText, StringComparison.Ordinal))
                    throw new CodecException(
                        "invalid_document_content_control",
                        $"Document drop-down content control {control.Id} visible text does not match its selected value.");
            }
        }
    }

    internal static bool IsSupported(W.SdtRun source)
    {
        var properties = source.SdtProperties;
        var content = source.SdtContentRun;
        if (properties is null || content is null || source.ChildElements.Count != 2) return false;
        if (properties.Elements<W.SdtAlias>().Count() > 1 ||
            properties.Elements<W.Tag>().Count() != 1 ||
            properties.Elements<W.SdtId>().Count() != 1) return false;
        var textCount = properties.Elements<W.SdtContentText>().Count();
        var checkboxCount = properties.Elements<W14.SdtContentCheckBox>().Count();
        var dropdownCount = properties.Elements<W.SdtContentDropDownList>().Count();
        if (textCount + checkboxCount + dropdownCount != 1) return false;
        if (properties.ChildElements.Any(child => child is not W.SdtAlias and not W.Tag and not W.SdtId and not W.SdtContentText and not W14.SdtContentCheckBox and not W.SdtContentDropDownList)) return false;
        var tag = properties.GetFirstChild<W.Tag>()?.Val?.Value;
        var alias = properties.GetFirstChild<W.SdtAlias>()?.Val?.Value ?? tag;
        var nativeId = properties.GetFirstChild<W.SdtId>()?.Val?.Value;
        if (!ValidText(tag, 1, MaxTagLength) || !ValidText(alias, 0, MaxAliasLength) || nativeId is null or <= 0) return false;
        if (textCount == 1 && properties.GetFirstChild<W.SdtContentText>()?.MultiLine?.Value == true) return false;
        if (content.ChildElements.Count != 1 || content.FirstChild is not W.Run run) return false;
        if (!run.ChildElements.All(child => child is W.RunProperties or W.Text) ||
            !DocxFormattingCodec.IsSupportedRunProperties(run.RunProperties)) return false;
        var visibleText = string.Concat(run.Elements<W.Text>().Select(text => text.Text));
        if (run.Elements<W.Text>().Count() != 1) return false;
        if (checkboxCount == 1)
        {
            var checkbox = properties.GetFirstChild<W14.SdtContentCheckBox>()!;
            if (checkbox.ChildElements.Count != 3 ||
                checkbox.Elements<W14.Checked>().Count() != 1 ||
                checkbox.Elements<W14.CheckedState>().Count() != 1 ||
                checkbox.Elements<W14.UncheckedState>().Count() != 1 ||
                checkbox.ChildElements.Any(child => child is not W14.Checked and not W14.CheckedState and not W14.UncheckedState)) return false;
            var checkedValue = checkbox.GetFirstChild<W14.Checked>()?.Val?.Value;
            var isChecked = checkedValue == W14.OnOffValues.One || checkedValue == W14.OnOffValues.True;
            var checkedState = checkbox.GetFirstChild<W14.CheckedState>();
            var uncheckedState = checkbox.GetFirstChild<W14.UncheckedState>();
            return checkedValue is not null &&
                   CheckedSymbolCode.Equals(checkedState?.Val?.Value, StringComparison.OrdinalIgnoreCase) &&
                   CheckboxFont.Equals(checkedState?.Font?.Value, StringComparison.Ordinal) &&
                   UncheckedSymbolCode.Equals(uncheckedState?.Val?.Value, StringComparison.OrdinalIgnoreCase) &&
                   CheckboxFont.Equals(uncheckedState?.Font?.Value, StringComparison.Ordinal) &&
                   visibleText.Equals(isChecked ? CheckedGlyph : UncheckedGlyph, StringComparison.Ordinal);
        }
        if (dropdownCount == 0) return true;
        var dropdown = properties.GetFirstChild<W.SdtContentDropDownList>()!;
        var choices = dropdown.Elements<W.ListItem>().ToArray();
        var lastValue = dropdown.LastValue?.Value;
        if (dropdown.ExtendedAttributes.Any() || choices.Length is < 1 or > MaxChoiceCount || choices.Length != dropdown.ChildElements.Count || !ValidText(lastValue, 1, MaxChoiceTextLength)) return false;
        var values = new HashSet<string>(StringComparer.Ordinal);
        var displayTexts = new HashSet<string>(StringComparer.Ordinal);
        string? selectedDisplayText = null;
        foreach (var choice in choices)
        {
            var displayText = choice.DisplayText?.Value;
            var value = choice.Value?.Value;
            if (choice.HasChildren || choice.ExtendedAttributes.Any() || !ValidText(displayText, 1, MaxChoiceTextLength) || !ValidText(value, 1, MaxChoiceTextLength) || !values.Add(value!) || !displayTexts.Add(displayText!)) return false;
            if (value!.Equals(lastValue, StringComparison.Ordinal)) selectedDisplayText = displayText;
        }
        return selectedDisplayText is not null && visibleText.Equals(selectedDisplayText, StringComparison.Ordinal);
    }

    internal static DocumentRun Read(W.SdtRun source, string modelId)
    {
        if (!IsSupported(source))
            throw new CodecException(
                "unsupported_document_content_control",
                "DOCX inline content control is outside the bounded plain-text, canonical checkbox, or canonical drop-down SDT profiles.",
                "word/document.xml");
        var properties = source.SdtProperties!;
        var result = DocxCodec.ReadRun((W.Run)source.SdtContentRun!.FirstChild!);
        var checkbox = properties.GetFirstChild<W14.SdtContentCheckBox>();
        var dropdown = properties.GetFirstChild<W.SdtContentDropDownList>();
        result.TextContentControl = new DocumentTextContentControl
        {
            Id = modelId,
            Tag = properties.GetFirstChild<W.Tag>()!.Val!.Value!,
            Alias = properties.GetFirstChild<W.SdtAlias>()?.Val?.Value ?? properties.GetFirstChild<W.Tag>()!.Val!.Value!,
            NativeId = checked((uint)properties.GetFirstChild<W.SdtId>()!.Val!.Value),
            ControlType = checkbox is not null
                ? DocumentContentControlType.Checkbox
                : dropdown is not null ? DocumentContentControlType.DropDown : DocumentContentControlType.PlainText,
            Checked = checkbox is not null && IsChecked(checkbox.GetFirstChild<W14.Checked>()?.Val?.Value),
            SelectedValue = dropdown?.LastValue?.Value ?? string.Empty,
        };
        if (dropdown is not null)
            result.TextContentControl.Choices.Add(dropdown.Elements<W.ListItem>().Select(choice => new DocumentContentControlChoice
            {
                DisplayText = choice.DisplayText!.Value!,
                Value = choice.Value!.Value!,
            }));
        return result;
    }

    internal static W.SdtRun Build(DocumentRun source)
    {
        var control = source.TextContentControl ?? throw new CodecException(
            "invalid_document_content_control",
            "Document content-control run has no control metadata.");
        var properties = new W.SdtProperties();
        if (control.Alias.Length > 0) properties.Append(new W.SdtAlias { Val = control.Alias });
        properties.Append(
            new W.Tag { Val = control.Tag },
            new W.SdtId { Val = checked((int)control.NativeId) });
        var controlType = NormalizedType(control);
        if (controlType == DocumentContentControlType.Checkbox)
            properties.Append(new W14.SdtContentCheckBox(
                new W14.Checked { Val = control.Checked ? W14.OnOffValues.One : W14.OnOffValues.Zero },
                new W14.CheckedState { Val = CheckedSymbolCode, Font = CheckboxFont },
                new W14.UncheckedState { Val = UncheckedSymbolCode, Font = CheckboxFont }));
        else if (controlType == DocumentContentControlType.DropDown)
        {
            var dropdown = new W.SdtContentDropDownList { LastValue = control.SelectedValue };
            dropdown.Append(control.Choices.Select(choice => new W.ListItem
            {
                DisplayText = choice.DisplayText,
                Value = choice.Value,
            }));
            properties.Append(dropdown);
        }
        else
            properties.Append(new W.SdtContentText());
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
                $"Imported document paragraph {blockId} content-control topology is source-bound.",
                "word/document.xml");
    }

    private static List<(int RunIndex, uint NativeId, DocumentContentControlType ControlType, string ChoiceSignature)> Topology(DocumentParagraph paragraph) =>
        paragraph.Runs
            .Select((run, index) => (Run: run, Index: index))
            .Where(item => item.Run.TextContentControl is not null)
            .Select(item => (item.Index, item.Run.TextContentControl.NativeId, NormalizedType(item.Run.TextContentControl), ChoiceSignature(item.Run.TextContentControl)))
            .ToList();

    private static IEnumerable<(DocumentBlock Block, DocumentRun Run, DocumentTextContentControl Control)> Controls(DocumentArtifact document) =>
        document.Blocks
            .Where(block => block.ContentCase == DocumentBlock.ContentOneofCase.Paragraph)
            .SelectMany(block => block.Paragraph.Runs
                .Where(run => run.TextContentControl is not null)
                .Select(run => (block, run, run.TextContentControl)));

    private static DocumentContentControlType NormalizedType(DocumentTextContentControl control) =>
        control.ControlType switch
        {
            DocumentContentControlType.Unspecified => DocumentContentControlType.PlainText,
            DocumentContentControlType.PlainText => DocumentContentControlType.PlainText,
            DocumentContentControlType.Checkbox => DocumentContentControlType.Checkbox,
            DocumentContentControlType.DropDown => DocumentContentControlType.DropDown,
            _ => throw new CodecException(
                "invalid_document_content_control",
                $"Document content control {control.Id} has an unsupported control type."),
        };

    private static DocumentContentControlChoice ValidateDropdown(DocumentTextContentControl control)
    {
        if (control.Choices.Count is < 1 or > MaxChoiceCount)
            throw new CodecException(
                "invalid_document_content_control",
                $"Document drop-down content control {control.Id} requires 1 through {MaxChoiceCount} choices.");
        var values = new HashSet<string>(StringComparer.Ordinal);
        var displayTexts = new HashSet<string>(StringComparer.Ordinal);
        DocumentContentControlChoice? selected = null;
        foreach (var choice in control.Choices)
        {
            ValidateText(choice.DisplayText, $"Document drop-down content control {control.Id} choice display text", 1, MaxChoiceTextLength);
            ValidateText(choice.Value, $"Document drop-down content control {control.Id} choice value", 1, MaxChoiceTextLength);
            if (!values.Add(choice.Value) || !displayTexts.Add(choice.DisplayText))
                throw new CodecException(
                    "invalid_document_content_control",
                    $"Document drop-down content control {control.Id} choice values and display text must be unique.");
            if (choice.Value.Equals(control.SelectedValue, StringComparison.Ordinal)) selected = choice;
        }
        return selected ?? throw new CodecException(
            "invalid_document_content_control",
            $"Document drop-down content control {control.Id} selected value must match one declared choice.");
    }

    private static string ChoiceSignature(DocumentTextContentControl control) =>
        NormalizedType(control) == DocumentContentControlType.DropDown
            ? string.Join(".", control.Choices.Select(choice =>
                $"{Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(choice.DisplayText))}:{Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(choice.Value))}"))
            : string.Empty;

    private static bool IsChecked(W14.OnOffValues? value) =>
        value == W14.OnOffValues.One || value == W14.OnOffValues.True;

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
