using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenChestnut.Codec;

// Owns the narrow SpreadsheetML number-format slice. Existing cell formats are
// cloned instead of mutated so unrelated font, fill, border, alignment,
// protection, extension, and named-style semantics remain attached to a cell.
internal sealed class XlsxNumberFormatCodec
{
    internal const int MaxFormatCodeLength = 4_096;

    private static readonly IReadOnlyDictionary<uint, string> BuiltInFormats = new Dictionary<uint, string>
    {
        [0] = string.Empty,
        [1] = "0",
        [2] = "0.00",
        [3] = "#,##0",
        [4] = "#,##0.00",
        [5] = "\"$\"#,##0_);(\"$\"#,##0)",
        [6] = "\"$\"#,##0_);[Red](\"$\"#,##0)",
        [7] = "\"$\"#,##0.00_);(\"$\"#,##0.00)",
        [8] = "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)",
        [9] = "0%",
        [10] = "0.00%",
        [11] = "0.00E+00",
        [12] = "# ?/?",
        [13] = "# ??/??",
        [14] = "mm-dd-yy",
        [15] = "d-mmm-yy",
        [16] = "d-mmm",
        [17] = "mmm-yy",
        [18] = "h:mm AM/PM",
        [19] = "h:mm:ss AM/PM",
        [20] = "h:mm",
        [21] = "h:mm:ss",
        [22] = "m/d/yy h:mm",
        [37] = "#,##0_);(#,##0)",
        [38] = "#,##0_);[Red](#,##0)",
        [39] = "#,##0.00_);(#,##0.00)",
        [40] = "#,##0.00_);[Red](#,##0.00)",
        [41] = "_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"_);_(@_)",
        [42] = "_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)",
        [43] = "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)",
        [44] = "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)",
        [45] = "mm:ss",
        [46] = "[h]:mm:ss",
        [47] = "mmss.0",
        [48] = "##0.0E+0",
        [49] = "@",
    };

    private static readonly IReadOnlyDictionary<string, uint> BuiltInIds = BuiltInFormats
        .Where(item => item.Key != 0)
        .GroupBy(item => item.Value, StringComparer.Ordinal)
        .ToDictionary(group => group.Key, group => group.First().Key, StringComparer.Ordinal);

    private readonly WorkbookPart _workbookPart;
    private WorkbookStylesPart? _stylesPart;
    private Stylesheet? _stylesheet;
    private readonly Dictionary<uint, string> _customFormatsById = [];
    private readonly Dictionary<string, uint> _customFormatIds = new(StringComparer.Ordinal);
    private readonly Dictionary<(uint SourceStyleIndex, string FormatCode), uint> _derivedStyleIndexes = [];
    private List<CellFormat> _cellFormats = [];
    private string[] _originalNumberingFormats = [];
    private string[] _originalCellFormats = [];
    private string[] _originalOtherChildren = [];
    private bool _dirty;

    internal XlsxNumberFormatCodec(WorkbookPart workbookPart)
    {
        _workbookPart = workbookPart;
        _stylesPart = workbookPart.WorkbookStylesPart;
        _stylesheet = _stylesPart?.Stylesheet;
        if (_stylesPart is not null && _stylesheet is null)
            throw Invalid("Workbook styles part has no stylesheet root.");
        if (_stylesheet is not null) IndexExistingStyles(_stylesheet);
    }

    internal string Read(Cell cell)
    {
        if (cell.StyleIndex?.HasValue != true) return string.Empty;
        if (_stylesheet?.CellFormats is null)
            throw Invalid($"Cell {cell.CellReference} references style {cell.StyleIndex.Value}, but the workbook has no cell formats.");
        var styleIndex = cell.StyleIndex.Value;
        if (styleIndex >= _cellFormats.Count)
            throw Invalid($"Cell {cell.CellReference} references missing style {styleIndex}.");
        return ResolveFormatCode(_cellFormats[checked((int)styleIndex)]);
    }

    internal void Apply(Cell cell, string? formatCode)
    {
        var desired = Canonicalize(formatCode, cell.CellReference?.Value);
        var current = Read(cell);
        if (string.Equals(current, desired, StringComparison.Ordinal)) return;

        EnsureWritableStylesheet();
        var sourceStyleIndex = cell.StyleIndex?.Value ?? 0;
        if (sourceStyleIndex >= _cellFormats.Count)
            throw Invalid($"Cell {cell.CellReference} references missing style {sourceStyleIndex}.");
        if (_derivedStyleIndexes.TryGetValue((sourceStyleIndex, desired), out var existingStyleIndex))
        {
            cell.StyleIndex = existingStyleIndex;
            return;
        }

        var derived = (CellFormat)_cellFormats[checked((int)sourceStyleIndex)].CloneNode(true);
        derived.NumberFormatId = FindOrCreateFormatId(desired);
        derived.ApplyNumberFormat = true;
        _stylesheet!.CellFormats!.Append(derived);
        _cellFormats.Add(derived);
        _stylesheet.CellFormats.Count = checked((uint)_cellFormats.Count);
        var styleIndex = checked((uint)_cellFormats.Count - 1);
        _derivedStyleIndexes[(sourceStyleIndex, desired)] = styleIndex;
        cell.StyleIndex = styleIndex;
        _dirty = true;
    }

    internal void Save()
    {
        if (!_dirty || _stylesheet is null) return;
        AssertOriginalStylesPreserved();
        _stylesheet.Save();
    }

    internal static string Canonicalize(string? formatCode, string? sourceIdentity = null)
    {
        var value = formatCode ?? string.Empty;
        if (value.Equals("General", StringComparison.OrdinalIgnoreCase)) return string.Empty;
        if (value.Length > MaxFormatCodeLength)
            throw Invalid($"Number format for {sourceIdentity ?? "cell"} exceeds {MaxFormatCodeLength} characters.");
        if (value.Any(character => char.IsControl(character)))
            throw Invalid($"Number format for {sourceIdentity ?? "cell"} contains a control character.");
        return value;
    }

    private string ResolveFormatCode(CellFormat cellFormat)
    {
        var numberFormatId = cellFormat.NumberFormatId?.Value ?? 0;
        if (cellFormat.ApplyNumberFormat?.HasValue == true && cellFormat.ApplyNumberFormat.Value == false)
            numberFormatId = ResolveBaseFormatId(cellFormat);
        else if (numberFormatId == 0)
            numberFormatId = ResolveBaseFormatId(cellFormat);

        if (BuiltInFormats.TryGetValue(numberFormatId, out var builtIn)) return builtIn;
        if (_customFormatsById.TryGetValue(numberFormatId, out var custom)) return custom;
        if (numberFormatId < 164)
            throw Invalid($"Workbook uses locale-dependent built-in number format {numberFormatId}, which this codec cannot represent as a stable format code.");
        throw Invalid($"Workbook references missing custom number format {numberFormatId}.");
    }

    private uint ResolveBaseFormatId(CellFormat cellFormat)
    {
        var formatId = cellFormat.FormatId?.Value;
        if (formatId is null || _stylesheet?.CellStyleFormats is null) return 0;
        var bases = _stylesheet.CellStyleFormats.Elements<CellFormat>().ToArray();
        if (formatId.Value >= bases.Length)
            throw Invalid($"Cell format references missing base style {formatId.Value}.");
        return bases[checked((int)formatId.Value)].NumberFormatId?.Value ?? 0;
    }

    private uint FindOrCreateFormatId(string formatCode)
    {
        if (formatCode.Length == 0) return 0;
        if (BuiltInIds.TryGetValue(formatCode, out var builtInId)) return builtInId;
        if (_customFormatIds.TryGetValue(formatCode, out var existingId)) return existingId;

        var nextId = Math.Max(164U, _customFormatsById.Keys.DefaultIfEmpty(163U).Max() + 1U);
        var numberingFormats = _stylesheet!.NumberingFormats;
        if (numberingFormats is null)
        {
            numberingFormats = new NumberingFormats();
            _stylesheet.InsertAt(numberingFormats, 0);
        }
        numberingFormats.Append(new NumberingFormat { NumberFormatId = nextId, FormatCode = formatCode });
        numberingFormats.Count = checked((uint)numberingFormats.ChildElements.Count);
        _customFormatsById[nextId] = formatCode;
        _customFormatIds[formatCode] = nextId;
        _dirty = true;
        return nextId;
    }

    private void EnsureWritableStylesheet()
    {
        if (_stylesheet is not null)
        {
            if (_stylesheet.CellFormats is null || _cellFormats.Count == 0)
                throw Invalid("Workbook stylesheet has no default cell format.");
            return;
        }

        _stylesPart = _workbookPart.AddNewPart<WorkbookStylesPart>();
        _stylesheet = CreateMinimalStylesheet();
        _stylesPart.Stylesheet = _stylesheet;
        IndexExistingStyles(_stylesheet);
        _dirty = true;
    }

    private void IndexExistingStyles(Stylesheet stylesheet)
    {
        _customFormatsById.Clear();
        _customFormatIds.Clear();
        foreach (var format in stylesheet.NumberingFormats?.Elements<NumberingFormat>() ?? [])
        {
            if (format.NumberFormatId?.HasValue != true || format.FormatCode?.HasValue != true)
                throw Invalid("Workbook contains an incomplete custom number-format definition.");
            var id = format.NumberFormatId.Value;
            var code = Canonicalize(format.FormatCode.Value, $"numFmt {id}");
            if (id < 164) throw Invalid($"Custom number format {id} uses the reserved built-in ID range.");
            if (!_customFormatsById.TryAdd(id, code)) throw Invalid($"Workbook defines custom number format {id} more than once.");
            _customFormatIds.TryAdd(code, id);
        }
        _cellFormats = stylesheet.CellFormats?.Elements<CellFormat>().ToList() ?? [];
        _originalNumberingFormats = stylesheet.NumberingFormats?.Elements<NumberingFormat>().Select(item => item.OuterXml).ToArray() ?? [];
        _originalCellFormats = _cellFormats.Select(item => item.OuterXml).ToArray();
        _originalOtherChildren = stylesheet.ChildElements
            .Where(item => item is not NumberingFormats and not CellFormats)
            .Select(item => item.OuterXml)
            .ToArray();
    }

    private void AssertOriginalStylesPreserved()
    {
        var numberingFormats = _stylesheet!.NumberingFormats?.Elements<NumberingFormat>().Select(item => item.OuterXml).ToArray() ?? [];
        var cellFormats = _stylesheet.CellFormats?.Elements<CellFormat>().Select(item => item.OuterXml).ToArray() ?? [];
        var otherChildren = _stylesheet.ChildElements
            .Where(item => item is not NumberingFormats and not CellFormats)
            .Select(item => item.OuterXml)
            .ToArray();
        if (!numberingFormats.Take(_originalNumberingFormats.Length).SequenceEqual(_originalNumberingFormats, StringComparer.Ordinal) ||
            !cellFormats.Take(_originalCellFormats.Length).SequenceEqual(_originalCellFormats, StringComparer.Ordinal) ||
            !otherChildren.SequenceEqual(_originalOtherChildren, StringComparer.Ordinal))
            throw new CodecException("style_preservation_failed", "Existing workbook style records changed while applying a cell number format.", "xl/styles.xml");
    }

    private static Stylesheet CreateMinimalStylesheet() => new(
        new Fonts(new Font()) { Count = 1U },
        new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 })) { Count = 2U },
        new Borders(new Border()) { Count = 1U },
        new CellStyleFormats(new CellFormat()) { Count = 1U },
        new CellFormats(new CellFormat()) { Count = 1U },
        new CellStyles(new CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U }) { Count = 1U });

    private static CodecException Invalid(string message) => new("invalid_cell_number_format", message, "xl/styles.xml");
}
