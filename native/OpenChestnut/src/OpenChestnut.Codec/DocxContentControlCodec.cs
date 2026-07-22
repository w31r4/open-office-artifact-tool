using System.Globalization;
using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Owns eleven deliberately bounded WordprocessingML SDT profiles/placements:
// inline and whole-table-cell plain-text, canonical drop-down, canonical
// combo-box, Word 2010+ checkbox, and ISO/Gregorian date controls, plus a block
// plain-text control. Each profile wraps exactly one paragraph containing one
// ordinary run and carries alias, tag, and native ID.
// Checkbox symbols plus list/date-control visible text are codec-owned. Every
// richer SDT remains opaque/source-bound.
internal static class DocxContentControlCodec
{
    private sealed record ControlReference(
        string Owner,
        DocumentBlock? Block,
        DocumentRun? Run,
        DocumentTextContentControl Control,
        bool IsBlock,
        string VisibleText)
    {
        internal bool IsTableCell => Block?.ContentCase == DocumentBlock.ContentOneofCase.Table;
    }

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
    private const string DateDisplayFormat = "yyyy-MM-dd";
    private const string DateLanguageId = "en-US";

    internal static void AssignNativeIds(DocumentArtifact document)
    {
        var controls = Controls(document).ToArray();
        var used = controls
            .Where(item => item.Control.HasNativeId && item.Control.NativeId is > 0 and <= int.MaxValue)
            .Select(item => item.Control.NativeId)
            .ToHashSet();
        uint next = 1;
        foreach (var item in controls)
        {
            var control = item.Control;
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
        foreach (var item in Controls(document))
        {
            var control = item.Control;
            ValidateText(control.Id, $"{item.Owner} content-control model ID", 1, 255);
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
            if (item.IsBlock)
            {
                var invalidParagraphBlock = !item.IsTableCell && item.Block is { } block &&
                    (block.Paragraph.Numbering is not null ||
                     block.Paragraph.Runs.Count != 1 ||
                     block.Paragraph.Runs[0].TextContentControl is not null ||
                     block.Paragraph.Runs[0].InlineField is not null ||
                     !block.Paragraph.Text.Equals(block.Paragraph.Runs[0].Text, StringComparison.Ordinal));
                if ((!item.IsTableCell && controlType != DocumentContentControlType.PlainText) || control.Alias.Length == 0 || invalidParagraphBlock)
                    throw new CodecException(
                        "invalid_document_content_control",
                        $"{item.Owner} content control {control.Id} must use the bounded placement/type profile around exactly one ordinary paragraph run.");
            }
            if (controlType == DocumentContentControlType.Checkbox &&
                !item.VisibleText.Equals(control.Checked ? CheckedGlyph : UncheckedGlyph, StringComparison.Ordinal))
                throw new CodecException(
                    "invalid_document_content_control",
                    $"Document checkbox content control {control.Id} visible glyph does not match its checked state.");
            if (controlType == DocumentContentControlType.DropDown)
            {
                var selected = ValidateDropdown(control);
                if (!item.VisibleText.Equals(selected.DisplayText, StringComparison.Ordinal))
                    throw new CodecException(
                        "invalid_document_content_control",
                        $"Document drop-down content control {control.Id} visible text does not match its selected value.");
            }
            if (controlType == DocumentContentControlType.ComboBox)
            {
                var visibleText = ValidateComboBox(control);
                if (!item.VisibleText.Equals(visibleText, StringComparison.Ordinal))
                    throw new CodecException(
                        "invalid_document_content_control",
                        $"Document combo-box content control {control.Id} visible text does not match its value.");
            }
            if (controlType == DocumentContentControlType.Date)
            {
                var dateValue = ValidateDateValue(control.DateValue, control.Id);
                if (!item.VisibleText.Equals(dateValue, StringComparison.Ordinal))
                    throw new CodecException(
                        "invalid_document_content_control",
                        $"Document date content control {control.Id} visible text does not match its date value.");
            }
        }
    }

    internal static bool IsSupported(W.SdtBlock source, bool allowTyped = false)
    {
        var properties = source.SdtProperties;
        var content = source.SdtContentBlock;
        if (properties is null || content is null || source.ChildElements.Count != 2 ||
            source.ChildElements[0] is not W.SdtProperties || source.ChildElements[1] is not W.SdtContentBlock ||
            source.GetAttributes().Count != 0 || properties.GetAttributes().Count != 0 || content.GetAttributes().Count != 0) return false;
        if (content.ChildElements.Count != 1 || content.FirstChild is not W.Paragraph paragraph ||
            !DocxFormattingCodec.IsSupportedParagraphProperties(paragraph.ParagraphProperties)) return false;
        var paragraphChildren = paragraph.ChildElements.Where(child => child is not W.ParagraphProperties).ToArray();
        if (paragraphChildren.Length != 1 || paragraphChildren[0] is not W.Run run ||
            run.Elements<W.RunProperties>().Count() > 1 || run.Elements<W.Text>().Count() != 1 ||
            run.ChildElements.Any(child => child is not W.RunProperties and not W.Text) ||
            !DocxFormattingCodec.IsSupportedRunProperties(run.RunProperties)) return false;
        var visibleText = string.Concat(run.Elements<W.Text>().Select(value => value.Text));
        if (!IsSupportedControlProperties(properties, visibleText, requireAlias: true, exactOrder: true) ||
            !allowTyped && properties.GetFirstChild<W.SdtContentText>() is null) return false;
        try
        {
            var modeledRun = DocxCodec.ReadRun(run);
            var modeled = new DocumentBlock
            {
                StyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? string.Empty,
                Paragraph = new DocumentParagraph
                {
                    Text = modeledRun.Text,
                    Formatting = DocxFormattingCodec.ReadParagraphFormatting(paragraph.ParagraphProperties),
                },
            };
            modeled.Paragraph.Runs.Add(modeledRun);
            return paragraph.OuterXml.Equals(DocxCodec.BuildParagraph(modeled).OuterXml, StringComparison.Ordinal);
        }
        catch (CodecException)
        {
            return false;
        }
    }

    internal static DocumentParagraph Read(W.SdtBlock source, string modelId, bool allowTyped = false)
    {
        if (!IsSupported(source, allowTyped))
            throw new CodecException(
                "unsupported_document_content_control",
                allowTyped
                    ? "DOCX table-cell content control is outside the bounded one-paragraph text, checkbox, drop-down, combo-box, or date SDT profiles."
                    : "DOCX block content control is outside the bounded one-paragraph plain-text SDT profile.",
                "word/document.xml");
        var properties = source.SdtProperties!;
        var paragraph = (W.Paragraph)source.SdtContentBlock!.FirstChild!;
        var run = (W.Run)paragraph.ChildElements.First(child => child is W.Run);
        var result = new DocumentParagraph
        {
            Text = string.Concat(run.Elements<W.Text>().Select(value => value.Text)),
            Formatting = DocxFormattingCodec.ReadParagraphFormatting(paragraph.ParagraphProperties),
            BlockContentControl = ReadControl(properties, modelId),
        };
        result.Runs.Add(DocxCodec.ReadRun(run));
        return result;
    }

    internal static bool IsSupported(W.SdtRun source)
    {
        var properties = source.SdtProperties;
        var content = source.SdtContentRun;
        if (properties is null || content is null || source.ChildElements.Count != 2) return false;
        if (content.ChildElements.Count != 1 || content.FirstChild is not W.Run run) return false;
        if (!run.ChildElements.All(child => child is W.RunProperties or W.Text) ||
            !DocxFormattingCodec.IsSupportedRunProperties(run.RunProperties)) return false;
        var visibleText = string.Concat(run.Elements<W.Text>().Select(text => text.Text));
        if (run.Elements<W.Text>().Count() != 1) return false;
        return IsSupportedControlProperties(properties, visibleText, requireAlias: false, exactOrder: false);
    }

    internal static DocumentRun Read(W.SdtRun source, string modelId)
    {
        if (!IsSupported(source))
            throw new CodecException(
                "unsupported_document_content_control",
                "DOCX inline content control is outside the bounded plain-text, canonical checkbox, canonical drop-down, canonical combo-box, or canonical date SDT profiles.",
                "word/document.xml");
        var properties = source.SdtProperties!;
        var result = DocxCodec.ReadRun((W.Run)source.SdtContentRun!.FirstChild!);
        result.TextContentControl = ReadControl(properties, modelId);
        return result;
    }

    private static bool IsSupportedControlProperties(
        W.SdtProperties properties,
        string visibleText,
        bool requireAlias,
        bool exactOrder)
    {
        if ((requireAlias ? properties.Elements<W.SdtAlias>().Count() != 1 : properties.Elements<W.SdtAlias>().Count() > 1) ||
            properties.Elements<W.Tag>().Count() != 1 ||
            properties.Elements<W.SdtId>().Count() != 1) return false;
        var textCount = properties.Elements<W.SdtContentText>().Count();
        var checkboxCount = properties.Elements<W14.SdtContentCheckBox>().Count();
        var dropdownCount = properties.Elements<W.SdtContentDropDownList>().Count();
        var comboBoxCount = properties.Elements<W.SdtContentComboBox>().Count();
        var dateCount = properties.Elements<W.SdtContentDate>().Count();
        if (textCount + checkboxCount + dropdownCount + comboBoxCount + dateCount != 1 ||
            properties.ChildElements.Any(child => child is not W.SdtAlias and not W.Tag and not W.SdtId and not W.SdtContentText and not W14.SdtContentCheckBox and not W.SdtContentDropDownList and not W.SdtContentComboBox and not W.SdtContentDate)) return false;
        if (exactOrder &&
            (properties.ChildElements.Count != 4 ||
             properties.ChildElements[0] is not W.SdtAlias ||
             properties.ChildElements[1] is not W.Tag ||
             properties.ChildElements[2] is not W.SdtId ||
             !IsControlTypeProperty(properties.ChildElements[3]))) return false;
        var aliasElement = properties.GetFirstChild<W.SdtAlias>();
        var tagElement = properties.GetFirstChild<W.Tag>();
        var idElement = properties.GetFirstChild<W.SdtId>();
        if (exactOrder && new OpenXmlElement?[] { aliasElement, tagElement, idElement }
            .Any(child => child is null || child.HasChildren || child.ExtendedAttributes.Any())) return false;
        var tag = tagElement?.Val?.Value;
        var alias = aliasElement?.Val?.Value ?? tag;
        var nativeId = idElement?.Val?.Value;
        if (!ValidText(tag, 1, MaxTagLength) || !ValidText(alias, requireAlias ? 1 : 0, MaxAliasLength) || nativeId is null or <= 0) return false;
        if (textCount == 1)
        {
            var text = properties.GetFirstChild<W.SdtContentText>()!;
            return text.MultiLine?.Value != true && (!exactOrder || text.GetAttributes().Count == 0 && !text.HasChildren);
        }
        if (dateCount == 1) return IsCanonicalDate(properties.GetFirstChild<W.SdtContentDate>()!, visibleText);
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
        OpenXmlCompositeElement listControl;
        string? lastValue;
        if (dropdownCount == 1)
        {
            var dropdown = properties.GetFirstChild<W.SdtContentDropDownList>()!;
            listControl = dropdown;
            lastValue = dropdown.LastValue?.Value;
        }
        else
        {
            var comboBox = properties.GetFirstChild<W.SdtContentComboBox>()!;
            listControl = comboBox;
            lastValue = comboBox.LastValue?.Value;
        }
        var choices = listControl.Elements<W.ListItem>().ToArray();
        if (listControl.ExtendedAttributes.Any() || choices.Length is < 1 or > MaxChoiceCount || choices.Length != listControl.ChildElements.Count || !ValidText(lastValue, 1, MaxChoiceTextLength)) return false;
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
        return dropdownCount == 1
            ? selectedDisplayText is not null && visibleText.Equals(selectedDisplayText, StringComparison.Ordinal)
            : visibleText.Equals(selectedDisplayText ?? lastValue, StringComparison.Ordinal);
    }

    private static bool IsControlTypeProperty(OpenXmlElement element) =>
        element is W.SdtContentText or W14.SdtContentCheckBox or W.SdtContentDropDownList or W.SdtContentComboBox or W.SdtContentDate;

    private static DocumentTextContentControl ReadControl(W.SdtProperties properties, string modelId)
    {
        var checkbox = properties.GetFirstChild<W14.SdtContentCheckBox>();
        var dropdown = properties.GetFirstChild<W.SdtContentDropDownList>();
        var comboBox = properties.GetFirstChild<W.SdtContentComboBox>();
        var date = properties.GetFirstChild<W.SdtContentDate>();
        var result = new DocumentTextContentControl
        {
            Id = modelId,
            Tag = properties.GetFirstChild<W.Tag>()!.Val!.Value!,
            Alias = properties.GetFirstChild<W.SdtAlias>()?.Val?.Value ?? properties.GetFirstChild<W.Tag>()!.Val!.Value!,
            NativeId = checked((uint)properties.GetFirstChild<W.SdtId>()!.Val!.Value),
            ControlType = checkbox is not null
                ? DocumentContentControlType.Checkbox
                : dropdown is not null
                    ? DocumentContentControlType.DropDown
                    : comboBox is not null ? DocumentContentControlType.ComboBox
                        : date is not null ? DocumentContentControlType.Date : DocumentContentControlType.PlainText,
            Checked = checkbox is not null && IsChecked(checkbox.GetFirstChild<W14.Checked>()?.Val?.Value),
            SelectedValue = dropdown?.LastValue?.Value ?? string.Empty,
            Value = comboBox?.LastValue?.Value ?? string.Empty,
            DateValue = date is null ? string.Empty : DateValue(date),
        };
        var listControl = (OpenXmlCompositeElement?)dropdown ?? comboBox;
        if (listControl is not null)
            result.Choices.Add(listControl.Elements<W.ListItem>().Select(choice => new DocumentContentControlChoice
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
        var properties = BuildProperties(control);
        return new W.SdtRun(properties, new W.SdtContentRun(DocxCodec.BuildRun(source)));
    }

    private static W.SdtProperties BuildProperties(DocumentTextContentControl control)
    {
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
        else if (controlType == DocumentContentControlType.ComboBox)
        {
            var comboBox = new W.SdtContentComboBox { LastValue = control.Value };
            comboBox.Append(control.Choices.Select(choice => new W.ListItem
            {
                DisplayText = choice.DisplayText,
                Value = choice.Value,
            }));
            properties.Append(comboBox);
        }
        else if (controlType == DocumentContentControlType.Date)
        {
            var dateValue = ValidateDateValue(control.DateValue, control.Id);
            properties.Append(new W.SdtContentDate(
                new W.DateFormat { Val = DateDisplayFormat },
                new W.LanguageId { Val = DateLanguageId },
                new W.SdtDateMappingType { Val = W.DateFormatValues.Date },
                new W.Calendar { Val = W.CalendarValues.Gregorian })
            {
                FullDate = new DateTimeValue { InnerText = $"{dateValue}T00:00:00Z" },
            });
        }
        else
            properties.Append(new W.SdtContentText());
        return properties;
    }

    internal static W.SdtBlock Build(DocumentBlock source)
    {
        var control = source.Paragraph.BlockContentControl ?? throw new CodecException(
            "invalid_document_content_control",
            $"Document paragraph {source.Id} has no block content-control metadata.");
        return BuildBlock(control, DocxCodec.BuildParagraph(source));
    }

    internal static W.SdtBlock BuildBlock(DocumentTextContentControl control, W.Paragraph paragraph)
    {
        return new W.SdtBlock(BuildProperties(control), new W.SdtContentBlock(paragraph));
    }

    internal static void ApplyTableCell(
        W.TableCell cell,
        DocumentTableCell requested,
        DocumentTableCell source,
        string visibleText,
        string owner)
    {
        var requestedControl = requested.TextContentControl;
        var sourceControl = source.TextContentControl;
        if ((requestedControl is null) != (sourceControl is null) ||
            requestedControl is not null && sourceControl is not null &&
            (requestedControl.NativeId != sourceControl.NativeId ||
             NormalizedType(requestedControl) != NormalizedType(sourceControl) ||
             !ChoiceSignature(requestedControl).Equals(ChoiceSignature(sourceControl), StringComparison.Ordinal)))
            throw new CodecException(
                "document_content_control_topology_changed",
                $"{owner} content-control topology is source-bound.",
                "word/document.xml");
        if (requestedControl is null) return;
        var content = cell.ChildElements.Where(child => child is not W.TableCellProperties).ToArray();
        if (content.Length != 1 || content[0] is not W.SdtBlock sdt || !IsSupported(sdt, allowTyped: true))
            throw new CodecException(
                "unsupported_document_content_control",
                $"{owner} no longer matches a canonical table-cell text, checkbox, drop-down, combo-box, or date SDT profile.",
                "word/document.xml");
        sdt.SdtProperties!.GetFirstChild<W.SdtAlias>()!.Val = requestedControl.Alias;
        sdt.SdtProperties.GetFirstChild<W.Tag>()!.Val = requestedControl.Tag;
        ApplyMutableState(sdt.SdtProperties, requestedControl);
        var text = sdt.SdtContentBlock!.Descendants<W.Text>().Single();
        text.Text = visibleText;
        text.Space = visibleText.Length != visibleText.Trim().Length ? SpaceProcessingModeValues.Preserve : null;
        var observed = Read(sdt, requestedControl.Id, allowTyped: true);
        if (!observed.Text.Equals(visibleText, StringComparison.Ordinal) ||
            !observed.BlockContentControl.Equals(requestedControl))
            throw new CodecException(
                "document_semantics_not_applied",
                $"{owner} content-control edit did not round trip through the canonical native profile.",
                "word/document.xml");
    }

    private static void ApplyMutableState(W.SdtProperties properties, DocumentTextContentControl control)
    {
        switch (NormalizedType(control))
        {
            case DocumentContentControlType.Checkbox:
                properties.GetFirstChild<W14.SdtContentCheckBox>()!.GetFirstChild<W14.Checked>()!.Val =
                    control.Checked ? W14.OnOffValues.One : W14.OnOffValues.Zero;
                break;
            case DocumentContentControlType.DropDown:
                properties.GetFirstChild<W.SdtContentDropDownList>()!.LastValue = control.SelectedValue;
                break;
            case DocumentContentControlType.ComboBox:
                properties.GetFirstChild<W.SdtContentComboBox>()!.LastValue = control.Value;
                break;
            case DocumentContentControlType.Date:
                var dateValue = ValidateDateValue(control.DateValue, control.Id);
                properties.GetFirstChild<W.SdtContentDate>()!.FullDate =
                    new DateTimeValue { InnerText = $"{dateValue}T00:00:00Z" };
                break;
        }
    }

    internal static void MaskTableCellModeledState(W.SdtBlock control)
    {
        var properties = control.SdtProperties!;
        properties.GetFirstChild<W.SdtAlias>()!.Val = string.Empty;
        properties.GetFirstChild<W.Tag>()!.Val = string.Empty;
        if (properties.GetFirstChild<W14.SdtContentCheckBox>() is { } checkbox)
            checkbox.GetFirstChild<W14.Checked>()!.Val = W14.OnOffValues.Zero;
        else if (properties.GetFirstChild<W.SdtContentDropDownList>() is { } dropdown)
            dropdown.LastValue = string.Empty;
        else if (properties.GetFirstChild<W.SdtContentComboBox>() is { } comboBox)
            comboBox.LastValue = string.Empty;
        else if (properties.GetFirstChild<W.SdtContentDate>() is { } date)
            date.FullDate = new DateTimeValue { InnerText = "0001-01-01T00:00:00Z" };
    }

    internal static void AssertTopology(DocumentParagraph requested, DocumentParagraph original, string blockId)
    {
        var requestedBlock = requested.BlockContentControl;
        var originalBlock = original.BlockContentControl;
        if ((requestedBlock is null) != (originalBlock is null) ||
            requestedBlock is not null && originalBlock is not null &&
            (requestedBlock.NativeId != originalBlock.NativeId || NormalizedType(requestedBlock) != NormalizedType(originalBlock)))
            throw new CodecException(
                "document_content_control_topology_changed",
                $"Imported document paragraph {blockId} block content-control topology is source-bound.",
                "word/document.xml");
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

    private static IEnumerable<ControlReference> Controls(DocumentArtifact document)
    {
        foreach (var block in document.Blocks)
        {
            if (block.ContentCase == DocumentBlock.ContentOneofCase.Paragraph)
            {
                if (block.Paragraph.BlockContentControl is not null)
                    yield return new ControlReference($"Document block {block.Id}", block, null, block.Paragraph.BlockContentControl, true, block.Paragraph.Text);
                foreach (var run in block.Paragraph.Runs.Where(run => run.TextContentControl is not null))
                    yield return new ControlReference($"Document block {block.Id}", block, run, run.TextContentControl, false, run.Text);
                continue;
            }
            if (block.ContentCase != DocumentBlock.ContentOneofCase.Table) continue;
            for (var rowIndex = 0; rowIndex < block.Table.Rows.Count; rowIndex++)
            {
                var row = block.Table.Rows[rowIndex];
                for (var cellIndex = 0; cellIndex < row.RichCells.Count; cellIndex++)
                {
                    var control = row.RichCells[cellIndex].TextContentControl;
                    if (control is null) continue;
                    yield return new ControlReference(
                        $"Document table {block.Id} cell {rowIndex},{cellIndex}",
                        block,
                        null,
                        control,
                        true,
                        cellIndex < row.Cells.Count ? row.Cells[cellIndex] : string.Empty);
                }
            }
        }
    }

    private static DocumentContentControlType NormalizedType(DocumentTextContentControl control) =>
        control.ControlType switch
        {
            DocumentContentControlType.Unspecified => DocumentContentControlType.PlainText,
            DocumentContentControlType.PlainText => DocumentContentControlType.PlainText,
            DocumentContentControlType.Checkbox => DocumentContentControlType.Checkbox,
            DocumentContentControlType.DropDown => DocumentContentControlType.DropDown,
            DocumentContentControlType.ComboBox => DocumentContentControlType.ComboBox,
            DocumentContentControlType.Date => DocumentContentControlType.Date,
            _ => throw new CodecException(
                "invalid_document_content_control",
                $"Document content control {control.Id} has an unsupported control type."),
        };

    private static DocumentContentControlChoice ValidateDropdown(DocumentTextContentControl control)
    {
        var selected = ValidateChoices(control, "drop-down", control.SelectedValue);
        return selected ?? throw new CodecException(
            "invalid_document_content_control",
            $"Document drop-down content control {control.Id} selected value must match one declared choice.");
    }

    private static string ValidateComboBox(DocumentTextContentControl control)
    {
        ValidateText(control.Value, $"Document combo-box content control {control.Id} value", 1, MaxChoiceTextLength);
        var selected = ValidateChoices(control, "combo-box", control.Value);
        return selected?.DisplayText ?? control.Value;
    }

    private static DocumentContentControlChoice? ValidateChoices(DocumentTextContentControl control, string label, string currentValue)
    {
        if (control.Choices.Count is < 1 or > MaxChoiceCount)
            throw new CodecException(
                "invalid_document_content_control",
                $"Document {label} content control {control.Id} requires 1 through {MaxChoiceCount} choices.");
        var values = new HashSet<string>(StringComparer.Ordinal);
        var displayTexts = new HashSet<string>(StringComparer.Ordinal);
        DocumentContentControlChoice? selected = null;
        foreach (var choice in control.Choices)
        {
            ValidateText(choice.DisplayText, $"Document {label} content control {control.Id} choice display text", 1, MaxChoiceTextLength);
            ValidateText(choice.Value, $"Document {label} content control {control.Id} choice value", 1, MaxChoiceTextLength);
            if (!values.Add(choice.Value) || !displayTexts.Add(choice.DisplayText))
                throw new CodecException(
                    "invalid_document_content_control",
                    $"Document {label} content control {control.Id} choice values and display text must be unique.");
            if (choice.Value.Equals(currentValue, StringComparison.Ordinal)) selected = choice;
        }
        return selected;
    }

    private static string ChoiceSignature(DocumentTextContentControl control) =>
        NormalizedType(control) is DocumentContentControlType.DropDown or DocumentContentControlType.ComboBox
            ? string.Join(".", control.Choices.Select(choice =>
                $"{Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(choice.DisplayText))}:{Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(choice.Value))}"))
            : string.Empty;

    private static bool IsCanonicalDate(W.SdtContentDate date, string visibleText)
    {
        if (date.ExtendedAttributes.Any() || date.ChildElements.Count != 4 ||
            date.ChildElements[0] is not W.DateFormat format ||
            date.ChildElements[1] is not W.LanguageId language ||
            date.ChildElements[2] is not W.SdtDateMappingType mapping ||
            date.ChildElements[3] is not W.Calendar calendar ||
            date.ChildElements.Any(child => child.HasChildren || child.ExtendedAttributes.Any()) ||
            !DateDisplayFormat.Equals(format.Val?.Value, StringComparison.Ordinal) ||
            !DateLanguageId.Equals(language.Val?.Value, StringComparison.Ordinal) ||
            mapping.Val?.Value != W.DateFormatValues.Date ||
            calendar.Val?.Value != W.CalendarValues.Gregorian) return false;
        var dateValue = DateValue(date);
        return string.Equals(date.FullDate?.InnerText, $"{dateValue}T00:00:00Z", StringComparison.Ordinal) &&
               ValidDateValue(dateValue) && visibleText.Equals(dateValue, StringComparison.Ordinal);
    }

    private static string DateValue(W.SdtContentDate date) =>
        date.FullDate?.InnerText is { Length: >= 10 } fullDate ? fullDate[..10] : string.Empty;

    private static bool ValidDateValue(string? value) =>
        value is not null && DateOnly.TryParseExact(value, DateDisplayFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out _);

    private static string ValidateDateValue(string? value, string controlId)
    {
        if (!ValidDateValue(value))
            throw new CodecException(
                "invalid_document_content_control",
                $"Document date content control {controlId} date value must be a real Gregorian date in YYYY-MM-DD form.");
        return value!;
    }

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
