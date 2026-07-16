using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Google.Protobuf;
using Google.Protobuf.Collections;
using OpenOffice.Artifact.Wire.V1;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xne = DocumentFormat.OpenXml.Office.Excel;

namespace OpenChestnut.Codec;

// Office 2010 sparklines live inside one worksheet extension rather than a
// relationship-backed part. This codec owns only reversible rectangular
// groups; non-contiguous formulas and unknown x14 markup stay in the validated
// source worksheet and are never surfaced as editable semantics.
internal sealed class XlsxSparklineCodec
{
    internal const string ExtensionUri = "{05C60535-1F16-4FD2-B633-F4F36F0B64E0}";
    private const int MaxGroups = 16_384;
    private const int MaxSparklines = 1_048_576;
    private static readonly Regex A1Range = new(
        "^(?:(?:'(?<quoted>(?:[^']|'')+)'|(?<bare>[^'!]+))!)?\\$?(?<firstColumn>[A-Za-z]{1,3})\\$?(?<firstRow>[1-9][0-9]*)(?::\\$?(?<lastColumn>[A-Za-z]{1,3})\\$?(?<lastRow>[1-9][0-9]*))?$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly HashSet<string> GroupAttributes = new(StringComparer.Ordinal)
    {
        "manualMax", "manualMin", "lineWeight", "type", "dateAxis", "displayEmptyCellsAs", "markers", "high", "low",
        "first", "last", "negative", "displayXAxis", "displayHidden", "minAxisType", "maxAxisType", "rightToLeft",
    };

    private readonly HashSet<string> _dirtyPartPaths = new(StringComparer.OrdinalIgnoreCase);
    internal IReadOnlySet<string> DirtyPartPaths => _dirtyPartPaths;

    internal void Apply(
        WorksheetPart worksheetPart,
        string worksheetId,
        RepeatedField<SpreadsheetSparklineGroupArtifact> groups,
        bool sourceBound,
        string? originalWorksheetXmlSha256 = null)
    {
        _dirtyPartPaths.Clear();
        Validate(groups, worksheetId);
        var worksheet = worksheetPart.Worksheet ?? throw Invalid(worksheetId, "has no worksheet root.");
        if (!sourceBound)
        {
            if (groups.Count == 0) return;
            var extension = EnsureSourceFreeExtension(worksheet, worksheetId);
            extension.Append(new X14.SparklineGroups(groups.Select(BuildGroup)));
            _dirtyPartPaths.Add(PartPath(worksheetPart));
            return;
        }

        var worksheetHash = originalWorksheetXmlSha256 ?? WorksheetXmlSha256(worksheetPart);
        var records = Scan(worksheetPart, worksheetId, worksheetHash);
        if (records.Count != groups.Count)
            throw new CodecException("invalid_spreadsheet_sparkline_topology", $"Worksheet {worksheetId} source-bound sparkline group count cannot change from {records.Count} to {groups.Count}.", PartPath(worksheetPart));

        var replacements = new List<(SparklineRecord Record, X14.SparklineGroup Replacement)>();
        for (var index = 0; index < groups.Count; index++)
        {
            var target = groups[index];
            var record = records[index];
            var binding = target.Source ?? throw Invalid(worksheetId, $"group {target.Id} is missing its source binding.");
            if (!string.Equals(target.Id, record.Artifact.Id, StringComparison.Ordinal) ||
                !string.Equals(binding.PartPath, PartPath(worksheetPart), StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(binding.WorksheetXmlSha256, worksheetHash, StringComparison.OrdinalIgnoreCase) ||
                binding.ExtensionOrdinal != record.ExtensionOrdinal || binding.GroupOrdinal != record.GroupOrdinal ||
                !string.Equals(binding.GroupXmlSha256, Hash(record.Element.OuterXml), StringComparison.OrdinalIgnoreCase))
                throw new CodecException("invalid_spreadsheet_sparkline_source", $"Worksheet {worksheetId} sparkline group {target.Id} no longer matches its validated source locator.", PartPath(worksheetPart));

            var semanticHash = SemanticHash(target);
            if (string.Equals(semanticHash, binding.SemanticSha256, StringComparison.OrdinalIgnoreCase)) continue;
            if (!binding.Editable || record.Artifact.Source?.Editable != true)
                throw new CodecException("unsupported_spreadsheet_sparkline_edit", $"Worksheet {worksheetId} sparkline group {target.Id} is read-only because its native profile is outside the editable subset.", PartPath(worksheetPart));
            replacements.Add((record, BuildGroup(target)));
        }

        foreach (var replacement in replacements)
        {
            replacement.Record.Element.InsertAfterSelf(replacement.Replacement);
            replacement.Record.Element.Remove();
        }
        if (replacements.Count > 0) _dirtyPartPaths.Add(PartPath(worksheetPart));
    }

    internal IReadOnlyList<SpreadsheetSparklineGroupArtifact> Read(WorksheetPart worksheetPart, string worksheetId)
    {
        var hash = WorksheetXmlSha256(worksheetPart);
        return Scan(worksheetPart, worksheetId, hash).Select(record => record.Artifact).ToArray();
    }

    internal static string WorksheetXmlSha256(WorksheetPart part) =>
        Hash(part.Worksheet?.OuterXml ?? throw new CodecException("missing_worksheet_root", "Worksheet has no root element.", PartPath(part)));

    private static WorksheetExtension EnsureSourceFreeExtension(Worksheet worksheet, string worksheetId)
    {
        var extensions = worksheet.GetFirstChild<WorksheetExtensionList>();
        if (extensions is null)
        {
            extensions = new WorksheetExtensionList();
            worksheet.Append(extensions);
        }
        var matching = extensions.Elements<WorksheetExtension>().Where(item => string.Equals(item.Uri?.Value, ExtensionUri, StringComparison.OrdinalIgnoreCase)).ToArray();
        if (matching.Length > 1 || matching.SingleOrDefault()?.ChildElements.Count > 0)
            throw Invalid(worksheetId, "has a conflicting sparkline extension in a source-free package.");
        if (matching.Length == 1) return matching[0];
        var extension = new WorksheetExtension { Uri = ExtensionUri };
        extensions.Append(extension);
        return extension;
    }

    private static List<SparklineRecord> Scan(WorksheetPart worksheetPart, string worksheetId, string worksheetHash)
    {
        var output = new List<SparklineRecord>();
        var extensions = worksheetPart.Worksheet?.GetFirstChild<WorksheetExtensionList>();
        if (extensions is null) return output;
        var partPath = PartPath(worksheetPart);
        for (var extensionOrdinal = 0; extensionOrdinal < extensions.ChildElements.Count; extensionOrdinal++)
        {
            if (extensions.ChildElements[extensionOrdinal] is not WorksheetExtension extension ||
                !string.Equals(extension.Uri?.Value, ExtensionUri, StringComparison.OrdinalIgnoreCase)) continue;
            var containers = extension.Elements<X14.SparklineGroups>().ToArray();
            if (containers.Length != 1) continue;
            var container = containers[0];
            for (var groupOrdinal = 0; groupOrdinal < container.ChildElements.Count; groupOrdinal++)
            {
                if (container.ChildElements[groupOrdinal] is not X14.SparklineGroup element ||
                    !TryReadGroup(element, out var artifact)) continue;
                artifact.Id = $"{worksheetId}/sparkline/{extensionOrdinal + 1}/{groupOrdinal + 1}";
                artifact.Source = new SpreadsheetSparklineSourceBinding
                {
                    PartPath = partPath,
                    WorksheetXmlSha256 = worksheetHash,
                    ExtensionOrdinal = checked((uint)extensionOrdinal),
                    GroupOrdinal = checked((uint)groupOrdinal),
                    GroupXmlSha256 = Hash(element.OuterXml),
                    Editable = true,
                };
                artifact.Source.SemanticSha256 = SemanticHash(artifact);
                output.Add(new SparklineRecord(artifact, element, checked((uint)extensionOrdinal), checked((uint)groupOrdinal)));
            }
        }
        return output;
    }

    private static bool TryReadGroup(X14.SparklineGroup source, out SpreadsheetSparklineGroupArtifact artifact)
    {
        artifact = new SpreadsheetSparklineGroupArtifact();
        if (source.GetAttributes().Any(attribute => !GroupAttributes.Contains(attribute.LocalName))) return false;
        if (source.Elements<X14.Sparklines>().SingleOrDefault() is not { } sparklines ||
            source.Elements<X14.Sparklines>().Skip(1).Any() ||
            source.ChildElements.Any(child => child is not X14.SeriesColor and not X14.NegativeColor and not X14.AxisColor and
                not X14.MarkersColor and not X14.FirstMarkerColor and not X14.LastMarkerColor and not X14.HighMarkerColor and
                not X14.LowMarkerColor and not Xne.Formula and not X14.Sparklines)) return false;
        if (!AtMostOne<X14.SeriesColor>(source) || !AtMostOne<X14.NegativeColor>(source) || !AtMostOne<X14.AxisColor>(source) ||
            !AtMostOne<X14.MarkersColor>(source) || !AtMostOne<X14.FirstMarkerColor>(source) || !AtMostOne<X14.LastMarkerColor>(source) ||
            !AtMostOne<X14.HighMarkerColor>(source) || !AtMostOne<X14.LowMarkerColor>(source) || !AtMostOne<Xne.Formula>(source)) return false;

        artifact.Type = source.Type?.Value switch
        {
            var value when value == X14.SparklineTypeValues.Line => SpreadsheetSparklineType.Line,
            var value when value == X14.SparklineTypeValues.Column => SpreadsheetSparklineType.Column,
            var value when value == X14.SparklineTypeValues.Stacked => SpreadsheetSparklineType.Stacked,
            null => SpreadsheetSparklineType.Line,
            _ => SpreadsheetSparklineType.Unspecified,
        };
        if (artifact.Type == SpreadsheetSparklineType.Unspecified) return false;
        artifact.LineWeight = source.LineWeight?.Value ?? 1D;
        artifact.DisplayHidden = source.DisplayHidden?.Value ?? false;
        artifact.DisplayEmptyCellsAs = source.DisplayEmptyCellsAs?.Value switch
        {
            var value when value == X14.DisplayBlanksAsValues.Span => SpreadsheetSparklineEmptyCells.Span,
            var value when value == X14.DisplayBlanksAsValues.Zero => SpreadsheetSparklineEmptyCells.Zero,
            var value when value == X14.DisplayBlanksAsValues.Gap => SpreadsheetSparklineEmptyCells.Gap,
            null => SpreadsheetSparklineEmptyCells.Gap,
            _ => SpreadsheetSparklineEmptyCells.Unspecified,
        };
        if (artifact.DisplayEmptyCellsAs == SpreadsheetSparklineEmptyCells.Unspecified) return false;
        artifact.Markers = new SpreadsheetSparklineMarkersArtifact
        {
            Show = source.Markers?.Value ?? false,
            High = source.High?.Value ?? false,
            Low = source.Low?.Value ?? false,
            First = source.First?.Value ?? false,
            Last = source.Last?.Value ?? false,
            Negative = source.Negative?.Value ?? false,
        };
        artifact.Axis = new SpreadsheetSparklineAxisArtifact
        {
            MinMode = AxisMode(source.MinAxisType?.Value),
            MaxMode = AxisMode(source.MaxAxisType?.Value),
            ShowAxis = source.DisplayXAxis?.Value ?? true,
            RightToLeft = source.RightToLeft?.Value ?? false,
        };
        if (source.ManualMin?.HasValue == true) artifact.Axis.ManualMin = source.ManualMin.Value;
        if (source.ManualMax?.HasValue == true) artifact.Axis.ManualMax = source.ManualMax.Value;
        if (artifact.Axis.MinMode == SpreadsheetSparklineAxisMode.Unspecified || artifact.Axis.MaxMode == SpreadsheetSparklineAxisMode.Unspecified) return false;

        if (!TryReadColor(source.SeriesColor, out var series) || !TryReadColor(source.NegativeColor, out var negative) ||
            !TryReadColor(source.AxisColor, out var axis) || !TryReadColor(source.MarkersColor, out var markers) ||
            !TryReadColor(source.FirstMarkerColor, out var first) || !TryReadColor(source.LastMarkerColor, out var last) ||
            !TryReadColor(source.HighMarkerColor, out var high) || !TryReadColor(source.LowMarkerColor, out var low)) return false;
        artifact.SeriesColor = series; artifact.NegativeColor = negative; artifact.AxisColor = axis; artifact.MarkersColor = markers;
        artifact.FirstMarkerColor = first; artifact.LastMarkerColor = last; artifact.HighMarkerColor = high; artifact.LowMarkerColor = low;

        var pairs = new List<SparklinePair>();
        foreach (var native in sparklines.Elements<X14.Sparkline>())
        {
            if (native.ChildElements.Count != 2 || native.Formula is null || native.ReferenceSequence is null ||
                !TryParseRange(native.Formula.Text, out var formula) || !TryParseRange(native.ReferenceSequence.Text, out var target) ||
                target.SheetName is not null || target.RowCount != 1 || target.ColumnCount != 1) return false;
            pairs.Add(new SparklinePair(formula, target));
        }
        if (pairs.Count == 0 || pairs.Count > MaxSparklines || !TryCollapsePairs(pairs, out var targetRange, out var sourceRange)) return false;
        artifact.TargetRange = targetRange;
        artifact.SourceDataRange = sourceRange;

        var dateFormula = source.Formula?.Text;
        var hasDateAxis = source.DateAxis?.Value ?? false;
        if (hasDateAxis != !string.IsNullOrWhiteSpace(dateFormula)) return false;
        if (hasDateAxis)
        {
            if (!TryParseRange(dateFormula!, out var dateRange) || dateRange.RowCount > 1 && dateRange.ColumnCount > 1 ||
                dateRange.CellCount != pairs[0].Formula.CellCount) return false;
            artifact.DateAxisRange = dateRange.QualifiedText;
        }
        artifact.Id = "imported-sparkline";
        try { ValidateArtifact(artifact, "imported worksheet"); }
        catch (CodecException) { return false; }
        return true;
    }

    private static X14.SparklineGroup BuildGroup(SpreadsheetSparklineGroupArtifact source)
    {
        var target = ParseRange(source.TargetRange, source.Id);
        var data = ParseRange(source.SourceDataRange, source.Id);
        var pairs = ExpandPairs(target, data, source.Id);
        var group = new X14.SparklineGroup
        {
            Type = source.Type switch
            {
                SpreadsheetSparklineType.Line => X14.SparklineTypeValues.Line,
                SpreadsheetSparklineType.Column => X14.SparklineTypeValues.Column,
                SpreadsheetSparklineType.Stacked => X14.SparklineTypeValues.Stacked,
                _ => throw Invalid(source.Id, "has an unsupported type."),
            },
            LineWeight = source.HasLineWeight ? source.LineWeight : null,
            DisplayHidden = source.HasDisplayHidden ? source.DisplayHidden : null,
            DisplayEmptyCellsAs = source.HasDisplayEmptyCellsAs ? source.DisplayEmptyCellsAs switch
            {
                SpreadsheetSparklineEmptyCells.Span => X14.DisplayBlanksAsValues.Span,
                SpreadsheetSparklineEmptyCells.Gap => X14.DisplayBlanksAsValues.Gap,
                SpreadsheetSparklineEmptyCells.Zero => X14.DisplayBlanksAsValues.Zero,
                _ => throw Invalid(source.Id, "has an unsupported empty-cell mode."),
            } : null,
        };
        var markers = source.Markers;
        if (markers is not null)
        {
            if (markers.HasShow) group.Markers = markers.Show;
            if (markers.HasHigh) group.High = markers.High;
            if (markers.HasLow) group.Low = markers.Low;
            if (markers.HasFirst) group.First = markers.First;
            if (markers.HasLast) group.Last = markers.Last;
            if (markers.HasNegative) group.Negative = markers.Negative;
        }
        var axis = source.Axis;
        if (axis is not null)
        {
            if (axis.HasManualMin) group.ManualMin = axis.ManualMin;
            if (axis.HasManualMax) group.ManualMax = axis.ManualMax;
            if (axis.HasMinMode) group.MinAxisType = NativeAxisMode(axis.MinMode, source.Id);
            if (axis.HasMaxMode) group.MaxAxisType = NativeAxisMode(axis.MaxMode, source.Id);
            if (axis.HasShowAxis) group.DisplayXAxis = axis.ShowAxis;
            if (axis.HasRightToLeft) group.RightToLeft = axis.RightToLeft;
        }
        AppendColor(group, source.SeriesColor, static () => new X14.SeriesColor());
        AppendColor(group, source.NegativeColor, static () => new X14.NegativeColor());
        AppendColor(group, source.AxisColor, static () => new X14.AxisColor());
        AppendColor(group, source.MarkersColor, static () => new X14.MarkersColor());
        AppendColor(group, source.FirstMarkerColor, static () => new X14.FirstMarkerColor());
        AppendColor(group, source.LastMarkerColor, static () => new X14.LastMarkerColor());
        AppendColor(group, source.HighMarkerColor, static () => new X14.HighMarkerColor());
        AppendColor(group, source.LowMarkerColor, static () => new X14.LowMarkerColor());
        if (source.DateAxisRange.Length > 0)
        {
            group.DateAxis = true;
            group.Append(new Xne.Formula(ParseRange(source.DateAxisRange, source.Id).QualifiedText));
        }
        var nativeSparklines = new X14.Sparklines();
        foreach (var pair in pairs)
            nativeSparklines.Append(new X14.Sparkline(new Xne.Formula(pair.Formula.QualifiedText), new Xne.ReferenceSequence(pair.Target.AddressText)));
        group.Append(nativeSparklines);
        return group;
    }

    private static void Validate(RepeatedField<SpreadsheetSparklineGroupArtifact> groups, string worksheetId)
    {
        if (groups.Count > MaxGroups) throw Invalid(worksheetId, $"exceeds the {MaxGroups}-group budget.");
        var ids = new HashSet<string>(StringComparer.Ordinal);
        var targets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var count = 0L;
        foreach (var group in groups)
        {
            ValidateArtifact(group, worksheetId);
            if (!ids.Add(group.Id)) throw Invalid(worksheetId, $"contains duplicate group ID {group.Id}.");
            var pairs = ExpandPairs(ParseRange(group.TargetRange, group.Id), ParseRange(group.SourceDataRange, group.Id), group.Id);
            count += pairs.Count;
            if (count > MaxSparklines) throw Invalid(worksheetId, $"exceeds the {MaxSparklines}-sparkline budget.");
            foreach (var pair in pairs)
                if (!targets.Add(pair.Target.AddressText)) throw Invalid(worksheetId, $"targets cell {pair.Target.AddressText} more than once.");
        }
    }

    private static void ValidateArtifact(SpreadsheetSparklineGroupArtifact group, string worksheetId)
    {
        if (string.IsNullOrWhiteSpace(group.Id) || group.Id.Length > 255 || group.Id.Any(char.IsControl)) throw Invalid(worksheetId, "contains a group with an invalid ID.");
        if (group.Type is not (SpreadsheetSparklineType.Line or SpreadsheetSparklineType.Column or SpreadsheetSparklineType.Stacked)) throw Invalid(worksheetId, $"group {group.Id} has an unsupported type.");
        if (!group.HasLineWeight || !double.IsFinite(group.LineWeight) || group.LineWeight <= 0 || group.LineWeight > 1584) throw Invalid(worksheetId, $"group {group.Id} line weight must be greater than 0 and no more than 1584 points.");
        if (!group.HasDisplayEmptyCellsAs || group.DisplayEmptyCellsAs is not (SpreadsheetSparklineEmptyCells.Span or SpreadsheetSparklineEmptyCells.Gap or SpreadsheetSparklineEmptyCells.Zero)) throw Invalid(worksheetId, $"group {group.Id} has an unsupported empty-cell mode.");
        var target = ParseRange(group.TargetRange, group.Id);
        var data = ParseRange(group.SourceDataRange, group.Id);
        var pairs = ExpandPairs(target, data, group.Id);
        if (group.DateAxisRange.Length > 0)
        {
            var date = ParseRange(group.DateAxisRange, group.Id);
            if (date.RowCount > 1 && date.ColumnCount > 1 || date.CellCount != pairs[0].Formula.CellCount) throw Invalid(worksheetId, $"group {group.Id} date axis must be one-dimensional and match each sparkline point count.");
        }
        var axis = group.Axis;
        if (axis is null || !axis.HasMinMode || !axis.HasMaxMode || axis.MinMode is SpreadsheetSparklineAxisMode.Unspecified || axis.MaxMode is SpreadsheetSparklineAxisMode.Unspecified) throw Invalid(worksheetId, $"group {group.Id} requires explicit min/max axis modes.");
        if (axis.HasManualMin && (!double.IsFinite(axis.ManualMin) || axis.MinMode != SpreadsheetSparklineAxisMode.Custom)) throw Invalid(worksheetId, $"group {group.Id} manual minimum requires a finite custom minimum mode.");
        if (axis.HasManualMax && (!double.IsFinite(axis.ManualMax) || axis.MaxMode != SpreadsheetSparklineAxisMode.Custom)) throw Invalid(worksheetId, $"group {group.Id} manual maximum requires a finite custom maximum mode.");
        if (axis.HasManualMin && axis.HasManualMax && axis.ManualMin >= axis.ManualMax) throw Invalid(worksheetId, $"group {group.Id} manual minimum must be less than its maximum.");
        foreach (var item in Colors(group)) ValidateColor(item.Color, worksheetId, group.Id, item.Name);
    }

    private static IReadOnlyList<SparklinePair> ExpandPairs(RangeRef target, RangeRef data, string id)
    {
        if (target.SheetName is not null) throw Invalid(id, "target range must not include a worksheet name.");
        if (target.RowCount > 1 && target.ColumnCount > 1) throw Invalid(id, "target range must be one-dimensional.");
        var output = new List<SparklinePair>(checked((int)target.CellCount));
        if (target.CellCount == 1)
        {
            if (data.RowCount > 1 && data.ColumnCount > 1) throw Invalid(id, "single-cell target requires one-dimensional source data.");
            output.Add(new SparklinePair(data, target));
            return output;
        }
        if (target.RowCount > 1)
        {
            if (data.RowCount != target.RowCount) throw Invalid(id, "vertical target count must match the source-data row count.");
            for (var index = 0; index < target.RowCount; index++) output.Add(new SparklinePair(
                data with { FirstRow = data.FirstRow + index, LastRow = data.FirstRow + index },
                target with { FirstRow = target.FirstRow + index, LastRow = target.FirstRow + index }));
            return output;
        }
        if (data.ColumnCount != target.ColumnCount) throw Invalid(id, "horizontal target count must match the source-data column count.");
        for (var index = 0; index < target.ColumnCount; index++) output.Add(new SparklinePair(
            data with { FirstColumn = data.FirstColumn + index, LastColumn = data.FirstColumn + index },
            target with { FirstColumn = target.FirstColumn + index, LastColumn = target.FirstColumn + index }));
        return output;
    }

    private static bool TryCollapsePairs(IReadOnlyList<SparklinePair> pairs, out string targetText, out string sourceText)
    {
        targetText = string.Empty; sourceText = string.Empty;
        if (pairs.Count == 1)
        {
            if (pairs[0].Formula.RowCount > 1 && pairs[0].Formula.ColumnCount > 1) return false;
            targetText = pairs[0].Target.AddressText;
            sourceText = pairs[0].Formula.QualifiedText;
            return true;
        }
        var vertical = pairs.All(pair => pair.Target.FirstColumn == pairs[0].Target.FirstColumn) &&
            pairs.Select((pair, index) => pair.Target.FirstRow == pairs[0].Target.FirstRow + index).All(value => value) &&
            pairs.All(pair => pair.Formula.SheetName == pairs[0].Formula.SheetName && pair.Formula.RowCount == 1 &&
                pair.Formula.FirstColumn == pairs[0].Formula.FirstColumn && pair.Formula.LastColumn == pairs[0].Formula.LastColumn) &&
            pairs.Select((pair, index) => pair.Formula.FirstRow == pairs[0].Formula.FirstRow + index).All(value => value);
        var horizontal = pairs.All(pair => pair.Target.FirstRow == pairs[0].Target.FirstRow) &&
            pairs.Select((pair, index) => pair.Target.FirstColumn == pairs[0].Target.FirstColumn + index).All(value => value) &&
            pairs.All(pair => pair.Formula.SheetName == pairs[0].Formula.SheetName && pair.Formula.ColumnCount == 1 &&
                pair.Formula.FirstRow == pairs[0].Formula.FirstRow && pair.Formula.LastRow == pairs[0].Formula.LastRow) &&
            pairs.Select((pair, index) => pair.Formula.FirstColumn == pairs[0].Formula.FirstColumn + index).All(value => value);
        if (!vertical && !horizontal) return false;
        var last = pairs[^1];
        var target = vertical
            ? pairs[0].Target with { LastRow = last.Target.LastRow }
            : pairs[0].Target with { LastColumn = last.Target.LastColumn };
        var data = vertical
            ? pairs[0].Formula with { LastRow = last.Formula.LastRow }
            : pairs[0].Formula with { LastColumn = last.Formula.LastColumn };
        targetText = target.AddressText;
        sourceText = data.QualifiedText;
        return true;
    }

    private static RangeRef ParseRange(string text, string location) =>
        TryParseRange(text, out var range) ? range : throw Invalid(location, $"contains invalid A1 range {text}.");

    private static bool TryParseRange(string text, out RangeRef range)
    {
        range = default!;
        if (string.IsNullOrWhiteSpace(text) || text.Length > 8_192) return false;
        var match = A1Range.Match(text.Trim());
        if (!match.Success) return false;
        var firstColumn = ColumnNumber(match.Groups["firstColumn"].Value);
        var lastColumn = match.Groups["lastColumn"].Success ? ColumnNumber(match.Groups["lastColumn"].Value) : firstColumn;
        if (!int.TryParse(match.Groups["firstRow"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var firstRow) ||
            !int.TryParse(match.Groups["lastRow"].Success ? match.Groups["lastRow"].Value : match.Groups["firstRow"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var lastRow) ||
            firstColumn < 1 || lastColumn < firstColumn || lastColumn > 16_384 || firstRow < 1 || lastRow < firstRow || lastRow > 1_048_576) return false;
        var sheetName = match.Groups["quoted"].Success ? match.Groups["quoted"].Value.Replace("''", "'", StringComparison.Ordinal)
            : match.Groups["bare"].Success ? match.Groups["bare"].Value : null;
        if (sheetName is { Length: > 31 } || sheetName?.Any(char.IsControl) == true) return false;
        range = new RangeRef(sheetName, firstRow, firstColumn, lastRow, lastColumn);
        return true;
    }

    private static int ColumnNumber(string text)
    {
        var value = 0;
        foreach (var character in text.ToUpperInvariant()) value = checked(value * 26 + character - 'A' + 1);
        return value;
    }

    private static string ColumnName(int number)
    {
        var builder = new StringBuilder();
        while (number > 0)
        {
            number--;
            builder.Insert(0, (char)('A' + number % 26));
            number /= 26;
        }
        return builder.ToString();
    }

    private static SpreadsheetSparklineAxisMode AxisMode(X14.SparklineAxisMinMaxValues? value)
    {
        if (value is null || value == X14.SparklineAxisMinMaxValues.Individual) return SpreadsheetSparklineAxisMode.Individual;
        if (value == X14.SparklineAxisMinMaxValues.Group) return SpreadsheetSparklineAxisMode.Group;
        if (value == X14.SparklineAxisMinMaxValues.Custom) return SpreadsheetSparklineAxisMode.Custom;
        return SpreadsheetSparklineAxisMode.Unspecified;
    }

    private static X14.SparklineAxisMinMaxValues NativeAxisMode(SpreadsheetSparklineAxisMode value, string id) => value switch
    {
        SpreadsheetSparklineAxisMode.Individual => X14.SparklineAxisMinMaxValues.Individual,
        SpreadsheetSparklineAxisMode.Group => X14.SparklineAxisMinMaxValues.Group,
        SpreadsheetSparklineAxisMode.Custom => X14.SparklineAxisMinMaxValues.Custom,
        _ => throw Invalid(id, "has an unsupported axis mode."),
    };

    private static bool TryReadColor(X14.ColorType? source, out SpreadsheetColor? color)
    {
        color = null;
        if (source is null) return true;
        var sources = (source.Rgb?.HasValue == true ? 1 : 0) + (source.Theme?.HasValue == true ? 1 : 0) +
            (source.Indexed?.HasValue == true ? 1 : 0) + (source.Auto?.Value == true ? 1 : 0);
        if (sources != 1) return false;
        var target = new SpreadsheetColor();
        if (source.Rgb?.HasValue == true)
        {
            var rgb = source.Rgb.Value ?? string.Empty;
            if (rgb.Length == 8 && rgb.StartsWith("FF", StringComparison.OrdinalIgnoreCase)) rgb = rgb[2..];
            if (rgb.Length != 6 || !rgb.All(Uri.IsHexDigit)) return false;
            target.Rgb = rgb.ToUpperInvariant();
        }
        else if (source.Theme?.HasValue == true && source.Theme.Value <= 11) target.Theme = source.Theme.Value;
        else if (source.Indexed?.HasValue == true && source.Indexed.Value <= 65) target.Indexed = source.Indexed.Value;
        else if (source.Auto?.Value == true) target.Automatic = true;
        else return false;
        if (source.Tint?.HasValue == true)
        {
            if (!double.IsFinite(source.Tint.Value) || source.Tint.Value is < -1 or > 1) return false;
            target.Tint = source.Tint.Value;
        }
        color = target;
        return true;
    }

    private static void AppendColor<T>(X14.SparklineGroup group, SpreadsheetColor? source, Func<T> factory) where T : X14.ColorType
    {
        if (source is null) return;
        var target = factory();
        target.Tint = source.HasTint ? source.Tint : null;
        switch (source.SourceCase)
        {
            case SpreadsheetColor.SourceOneofCase.Rgb: target.Rgb = $"FF{source.Rgb.ToUpperInvariant()}"; break;
            case SpreadsheetColor.SourceOneofCase.Theme: target.Theme = source.Theme; break;
            case SpreadsheetColor.SourceOneofCase.Indexed: target.Indexed = source.Indexed; break;
            case SpreadsheetColor.SourceOneofCase.Automatic: target.Auto = true; break;
            default: throw Invalid("sparkline", "color has no source.");
        }
        group.Append(target);
    }

    private static IEnumerable<(string Name, SpreadsheetColor? Color)> Colors(SpreadsheetSparklineGroupArtifact group)
    {
        yield return ("series", group.SeriesColor); yield return ("negative", group.NegativeColor); yield return ("axis", group.AxisColor);
        yield return ("markers", group.MarkersColor); yield return ("first marker", group.FirstMarkerColor); yield return ("last marker", group.LastMarkerColor);
        yield return ("high marker", group.HighMarkerColor); yield return ("low marker", group.LowMarkerColor);
    }

    private static void ValidateColor(SpreadsheetColor? color, string worksheetId, string groupId, string name)
    {
        if (color is null) return;
        if (color.SourceCase == SpreadsheetColor.SourceOneofCase.Rgb && (color.Rgb.Length != 6 || !color.Rgb.All(Uri.IsHexDigit))) throw Invalid(worksheetId, $"group {groupId} {name} color must be six-digit RGB.");
        if (color.SourceCase == SpreadsheetColor.SourceOneofCase.Theme && color.Theme > 11) throw Invalid(worksheetId, $"group {groupId} {name} theme index must be 0 through 11.");
        if (color.SourceCase == SpreadsheetColor.SourceOneofCase.Indexed && color.Indexed > 65) throw Invalid(worksheetId, $"group {groupId} {name} indexed color must be 0 through 65.");
        if (color.SourceCase == SpreadsheetColor.SourceOneofCase.Automatic && !color.Automatic) throw Invalid(worksheetId, $"group {groupId} {name} automatic color must be true.");
        if (color.SourceCase == SpreadsheetColor.SourceOneofCase.None) throw Invalid(worksheetId, $"group {groupId} {name} color has no source.");
        if (color.HasTint && (!double.IsFinite(color.Tint) || color.Tint is < -1 or > 1)) throw Invalid(worksheetId, $"group {groupId} {name} tint must be between -1 and 1.");
    }

    private static bool AtMostOne<T>(OpenXmlCompositeElement parent) where T : OpenXmlElement => !parent.Elements<T>().Skip(1).Any();

    private static string SemanticHash(SpreadsheetSparklineGroupArtifact source)
    {
        var semantic = source.Clone();
        semantic.Source = null;
        return Convert.ToHexString(SHA256.HashData(semantic.ToByteArray())).ToLowerInvariant();
    }

    private static string Hash(string text) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text))).ToLowerInvariant();
    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static CodecException Invalid(string worksheetId, string message) => new("invalid_spreadsheet_sparkline", $"Worksheet {worksheetId} sparkline {message}", worksheetId);

    private sealed record SparklineRecord(SpreadsheetSparklineGroupArtifact Artifact, X14.SparklineGroup Element, uint ExtensionOrdinal, uint GroupOrdinal);
    private sealed record SparklinePair(RangeRef Formula, RangeRef Target);
    private sealed record RangeRef(string? SheetName, int FirstRow, int FirstColumn, int LastRow, int LastColumn)
    {
        internal int RowCount => LastRow - FirstRow + 1;
        internal int ColumnCount => LastColumn - FirstColumn + 1;
        internal long CellCount => (long)RowCount * ColumnCount;
        internal string AddressText
        {
            get
            {
                var first = $"{ColumnName(FirstColumn)}{FirstRow}";
                var last = $"{ColumnName(LastColumn)}{LastRow}";
                return first == last ? first : $"{first}:{last}";
            }
        }
        internal string QualifiedText
        {
            get
            {
                if (SheetName is null) return AddressText;
                var sheet = Regex.IsMatch(SheetName, "^[A-Za-z_][A-Za-z0-9_.]*$", RegexOptions.CultureInvariant)
                    ? SheetName
                    : $"'{SheetName.Replace("'", "''", StringComparison.Ordinal)}'";
                return $"{sheet}!{AddressText}";
            }
        }
    }
}
