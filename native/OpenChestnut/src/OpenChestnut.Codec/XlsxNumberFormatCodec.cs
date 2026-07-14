namespace OpenChestnut.Codec;

// Pure number-format vocabulary shared by XlsxCellStyleCodec. Stylesheet
// ownership intentionally lives in one codec so a font/fill/border change and
// a number-format change produce one loss-aware derived cell-format record.
internal static class XlsxNumberFormatCodec
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

    internal static string Canonicalize(string? formatCode, string? sourceIdentity = null)
    {
        var value = formatCode ?? string.Empty;
        if (value.Equals("General", StringComparison.OrdinalIgnoreCase)) return string.Empty;
        if (value.Length > MaxFormatCodeLength)
            throw Invalid($"Number format for {sourceIdentity ?? "cell"} exceeds {MaxFormatCodeLength} characters.");
        if (value.Any(char.IsControl))
            throw Invalid($"Number format for {sourceIdentity ?? "cell"} contains a control character.");
        return value;
    }

    internal static bool TryGetBuiltInFormat(uint id, out string format) => BuiltInFormats.TryGetValue(id, out format!);

    internal static bool TryGetBuiltInId(string format, out uint id) => BuiltInIds.TryGetValue(format, out id);

    internal static CodecException Invalid(string message) => new("invalid_cell_number_format", message, "xl/styles.xml");
}
