using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Shared and legacy-array formulas are worksheet topology, not just cell text.
// This codec keeps that topology explicit and refuses malformed groups before
// package mutation so unrelated edits cannot silently flatten native formulas.
internal sealed class XlsxFormulaCodec
{
    private const int MaxFormulaLength = 8_192;
    private static readonly Regex CellReferencePattern = new(
        @"(?<![A-Za-z0-9_.])(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_. ]*))!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_])",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private readonly string _sheetName;
    private readonly Dictionary<(uint Row, uint Column), FormulaRecord> _records;

    private sealed record FormulaRecord(string Formula, CellFormulaMetadata? Metadata, CellFormula Source);
    private readonly record struct RangeBounds(uint Top, uint Left, uint Bottom, uint Right)
    {
        internal ulong CellCount => checked(((ulong)Bottom - Top + 1) * ((ulong)Right - Left + 1));
    }

    private XlsxFormulaCodec(string sheetName, Dictionary<(uint Row, uint Column), FormulaRecord> records)
    {
        _sheetName = sheetName;
        _records = records;
    }

    internal static XlsxFormulaCodec ForWorksheet(Worksheet worksheet, string sheetName)
    {
        var formulas = worksheet.GetFirstChild<SheetData>()?.Descendants<Cell>()
            .Where(cell => cell.CellFormula is not null)
            .Select(cell =>
            {
                var coordinate = ParseCellReference(cell.CellReference?.Value, sheetName);
                return (Coordinate: coordinate, Cell: cell, Formula: cell.CellFormula!);
            })
            .ToArray() ?? [];
        var records = new Dictionary<(uint Row, uint Column), FormulaRecord>();

        foreach (var entry in formulas.Where(item => FormulaType(item.Formula) != CellFormulaValues.Shared && FormulaType(item.Formula) != CellFormulaValues.Array))
        {
            var type = FormulaType(entry.Formula);
            if (type == CellFormulaValues.DataTable)
                throw Unsupported(entry.Cell, sheetName, "data-table formulas are outside the current OpenChestnut XLSX formula slice");
            if (type != CellFormulaValues.Normal)
                throw Unsupported(entry.Cell, sheetName, $"formula type {entry.Formula.FormulaType?.InnerText ?? "unknown"} is unsupported");
            var body = ValidateFormulaBody(entry.Formula.Text, entry.Cell.CellReference?.Value ?? "cell", required: true);
            records.Add(entry.Coordinate, new FormulaRecord($"={body}", null, entry.Formula));
        }

        foreach (var group in formulas.Where(item => FormulaType(item.Formula) == CellFormulaValues.Shared)
                     .GroupBy(item => item.Formula.SharedIndex?.Value ?? throw Invalid(item.Cell, sheetName, "shared formula is missing si")))
        {
            var masters = group.Where(item => !string.IsNullOrWhiteSpace(item.Formula.Text) || item.Formula.Reference?.Value is { Length: > 0 }).ToArray();
            if (masters.Length != 1)
                throw new CodecException("invalid_cell_formula", $"Worksheet {sheetName} shared formula si={group.Key} must have exactly one master; found {masters.Length}.", sheetName);
            var master = masters[0];
            if (string.IsNullOrWhiteSpace(master.Formula.Text) || master.Formula.Reference?.Value is not { Length: > 0 } reference)
                throw Invalid(master.Cell, sheetName, "shared formula master requires both formula text and ref");
            var bounds = ParseRange(reference, sheetName);
            if (master.Coordinate != (bounds.Top, bounds.Left))
                throw Invalid(master.Cell, sheetName, $"shared formula master must be the top-left cell of {reference}");
            var members = group.ToDictionary(item => item.Coordinate);
            if ((ulong)members.Count != bounds.CellCount)
                throw new CodecException("invalid_cell_formula", $"Worksheet {sheetName} shared formula si={group.Key} declares {reference} with {bounds.CellCount} cells but contains {members.Count} members.", sheetName);
            var masterBody = ValidateFormulaBody(master.Formula.Text, master.Cell.CellReference?.Value ?? "cell", required: true);
            for (var row = bounds.Top; row <= bounds.Bottom; row++)
            {
                for (var column = bounds.Left; column <= bounds.Right; column++)
                {
                    if (!members.TryGetValue((row, column), out var member))
                        throw new CodecException("invalid_cell_formula", $"Worksheet {sheetName} shared formula si={group.Key} is missing {CellReference(row, column)} from {reference}.", sheetName);
                    if (member.Coordinate != master.Coordinate && (!string.IsNullOrEmpty(member.Formula.Text) || member.Formula.Reference?.Value is { Length: > 0 }))
                        throw Invalid(member.Cell, sheetName, "shared formula follower must not contain formula text or ref");
                    var expanded = TranslateFormula(masterBody, master.Coordinate, member.Coordinate);
                    records.Add(member.Coordinate, new FormulaRecord($"={expanded}", new CellFormulaMetadata
                    {
                        Kind = CellFormulaKind.Shared,
                        SharedIndex = group.Key,
                        Reference = reference,
                    }, member.Formula));
                }
            }
        }

        foreach (var entry in formulas.Where(item => FormulaType(item.Formula) == CellFormulaValues.Array))
        {
            if (entry.Formula.Reference?.Value is not { Length: > 0 } reference)
                throw Invalid(entry.Cell, sheetName, "legacy array formula requires ref");
            var bounds = ParseRange(reference, sheetName);
            if (entry.Coordinate != (bounds.Top, bounds.Left))
                throw Invalid(entry.Cell, sheetName, $"legacy array formula must be anchored at the top-left cell of {reference}");
            var body = ValidateFormulaBody(entry.Formula.Text, entry.Cell.CellReference?.Value ?? "cell", required: true);
            records.Add(entry.Coordinate, new FormulaRecord($"={body}", new CellFormulaMetadata
            {
                Kind = CellFormulaKind.Array,
                Reference = reference,
            }, entry.Formula));
        }

        return new XlsxFormulaCodec(sheetName, records);
    }

    internal static void ValidateArtifact(WorksheetArtifact sheet)
    {
        var coordinates = new HashSet<(uint Row, uint Column)>();
        foreach (var cell in sheet.Cells)
        {
            if (!coordinates.Add((cell.Row, cell.Column)))
                throw new CodecException("duplicate_cell", $"Worksheet {sheet.Name} contains duplicate cell {CellReference(cell.Row, cell.Column)}.", sheet.Name);
            ValidateFormulaBody(cell.Formula, $"{sheet.Name}!{CellReference(cell.Row, cell.Column)}", required: cell.FormulaMetadata is not null);
            if (cell.FormulaMetadata is null) continue;
            if (cell.FormulaMetadata.Kind == CellFormulaKind.Unspecified)
                throw Invalid(cell, sheet.Name, "formula_metadata.kind must be shared or array");
            if (string.IsNullOrWhiteSpace(cell.FormulaMetadata.Reference))
                throw Invalid(cell, sheet.Name, "formula metadata requires reference");
        }

        foreach (var group in sheet.Cells.Where(cell => cell.FormulaMetadata?.Kind == CellFormulaKind.Shared)
                     .GroupBy(cell => cell.FormulaMetadata!.SharedIndex))
        {
            var references = group.Select(cell => cell.FormulaMetadata!.Reference).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            if (references.Length != 1)
                throw new CodecException("invalid_cell_formula", $"Worksheet {sheet.Name} shared formula si={group.Key} has inconsistent references.", sheet.Name);
            var reference = references[0];
            var bounds = ParseRange(reference, sheet.Name);
            var members = group.ToDictionary(cell => (cell.Row, cell.Column));
            if ((ulong)members.Count != bounds.CellCount)
                throw new CodecException("invalid_cell_formula", $"Worksheet {sheet.Name} shared formula si={group.Key} declares {reference} with {bounds.CellCount} cells but contains {members.Count} members.", sheet.Name);
            if (!members.TryGetValue((bounds.Top, bounds.Left), out var master))
                throw new CodecException("invalid_cell_formula", $"Worksheet {sheet.Name} shared formula si={group.Key} is missing master {CellReference(bounds.Top, bounds.Left)}.", sheet.Name);
            var masterBody = ValidateFormulaBody(master.Formula, $"{sheet.Name}!{CellReference(master.Row, master.Column)}", required: true);
            for (var row = bounds.Top; row <= bounds.Bottom; row++)
            {
                for (var column = bounds.Left; column <= bounds.Right; column++)
                {
                    if (!members.TryGetValue((row, column), out var member))
                        throw new CodecException("invalid_cell_formula", $"Worksheet {sheet.Name} shared formula si={group.Key} is missing {CellReference(row, column)} from {reference}.", sheet.Name);
                    var expected = $"={TranslateFormula(masterBody, (bounds.Top, bounds.Left), (row, column))}";
                    if (!string.Equals(NormalizeFormula(member.Formula), expected, StringComparison.Ordinal))
                        throw Invalid(member, sheet.Name, $"expanded shared formula must be {expected}");
                }
            }
        }

        var topology = new Dictionary<(uint Row, uint Column), string>();
        foreach (var cell in sheet.Cells.Where(cell => cell.FormulaMetadata is not null))
        {
            var metadata = cell.FormulaMetadata!;
            var bounds = ParseRange(metadata.Reference, sheet.Name);
            var owner = metadata.Kind == CellFormulaKind.Shared ? $"shared si={metadata.SharedIndex}" : $"array {CellReference(cell.Row, cell.Column)}";
            if (metadata.Kind == CellFormulaKind.Array && (cell.Row != bounds.Top || cell.Column != bounds.Left))
                throw Invalid(cell, sheet.Name, $"legacy array formula must be anchored at the top-left cell of {metadata.Reference}");
            if (metadata.Kind == CellFormulaKind.Array && metadata.SharedIndex != 0)
                throw Invalid(cell, sheet.Name, "legacy array formula must not set shared_index");
            for (var row = bounds.Top; row <= bounds.Bottom; row++)
            {
                for (var column = bounds.Left; column <= bounds.Right; column++)
                {
                    if (topology.TryGetValue((row, column), out var previous) && previous != owner)
                        throw Invalid(cell, sheet.Name, $"formula range {metadata.Reference} overlaps {previous}");
                    topology[(row, column)] = owner;
                }
            }
        }
    }

    internal void Populate(CellArtifact target)
    {
        if (!_records.TryGetValue((target.Row, target.Column), out var record)) return;
        target.Formula = record.Formula;
        if (record.Metadata is not null) target.FormulaMetadata = record.Metadata.Clone();
    }

    internal static CellFormula? Build(CellArtifact source)
    {
        if (string.IsNullOrWhiteSpace(source.Formula)) return null;
        var body = ValidateFormulaBody(source.Formula, CellReference(source.Row, source.Column), required: true);
        if (source.FormulaMetadata is null) return new CellFormula(body);
        var metadata = source.FormulaMetadata;
        var bounds = ParseRange(metadata.Reference, "worksheet");
        return metadata.Kind switch
        {
            CellFormulaKind.Shared when source.Row == bounds.Top && source.Column == bounds.Left => new CellFormula(body)
            {
                FormulaType = CellFormulaValues.Shared,
                SharedIndex = metadata.SharedIndex,
                Reference = metadata.Reference,
            },
            CellFormulaKind.Shared => new CellFormula
            {
                FormulaType = CellFormulaValues.Shared,
                SharedIndex = metadata.SharedIndex,
            },
            CellFormulaKind.Array => new CellFormula(body)
            {
                FormulaType = CellFormulaValues.Array,
                Reference = metadata.Reference,
            },
            _ => throw Invalid(source, "worksheet", "formula_metadata.kind must be shared or array"),
        };
    }

    internal void Apply(Cell target, CellArtifact desired)
    {
        var coordinate = (desired.Row, desired.Column);
        if (_records.TryGetValue(coordinate, out var current) && HasUnmodeledAttributes(current.Source))
            throw new CodecException("unsupported_cell_formula_edit", $"Cell {_sheetName}!{CellReference(desired.Row, desired.Column)} has unmodeled formula attributes and cannot be edited losslessly.", _sheetName);
        target.CellFormula = Build(desired);
    }

    internal static bool SemanticallyEqual(CellArtifact left, CellArtifact right)
    {
        if (!string.Equals(NormalizeFormula(left.Formula), NormalizeFormula(right.Formula), StringComparison.Ordinal)) return false;
        var leftMetadata = left.FormulaMetadata;
        var rightMetadata = right.FormulaMetadata;
        if (leftMetadata is null || rightMetadata is null) return leftMetadata is null && rightMetadata is null;
        return leftMetadata.Kind == rightMetadata.Kind && leftMetadata.SharedIndex == rightMetadata.SharedIndex &&
               string.Equals(leftMetadata.Reference, rightMetadata.Reference, StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasUnmodeledAttributes(CellFormula formula)
    {
        var type = FormulaType(formula);
        var allowed = type == CellFormulaValues.Shared
            ? new HashSet<string>(["t", "si", "ref"], StringComparer.Ordinal)
            : type == CellFormulaValues.Array
                ? new HashSet<string>(["t", "ref"], StringComparer.Ordinal)
                : new HashSet<string>(["t"], StringComparer.Ordinal);
        return formula.GetAttributes().Any(attribute => !allowed.Contains(attribute.LocalName));
    }

    private static CellFormulaValues FormulaType(CellFormula formula) => formula.FormulaType?.Value ?? CellFormulaValues.Normal;

    private static string ValidateFormulaBody(string? formula, string location, bool required)
    {
        var normalized = NormalizeFormula(formula);
        var body = normalized.StartsWith('=') ? normalized[1..] : normalized;
        if (required && string.IsNullOrWhiteSpace(body))
            throw new CodecException("invalid_cell_formula", $"Cell {location} requires non-empty formula text.", location);
        if (body.Length > MaxFormulaLength)
            throw new CodecException("invalid_cell_formula", $"Cell {location} formula exceeds {MaxFormulaLength} characters.", location);
        if (body.Any(char.IsControl))
            throw new CodecException("invalid_cell_formula", $"Cell {location} formula contains a control character.", location);
        return body;
    }

    private static string NormalizeFormula(string? formula)
    {
        var value = String(formula);
        if (value.Length == 0) return value;
        return value[0] == '=' ? value : $"={value}";
    }

    private static string String(string? value) => value ?? string.Empty;

    private static string TranslateFormula(string formula, (uint Row, uint Column) source, (uint Row, uint Column) target)
    {
        var rowOffset = checked((long)target.Row - source.Row);
        var columnOffset = checked((long)target.Column - source.Column);
        var output = new StringBuilder(formula.Length);
        var segment = new StringBuilder();
        var inString = false;
        var bracketDepth = 0;

        void Flush()
        {
            if (segment.Length == 0) return;
            output.Append(CellReferencePattern.Replace(segment.ToString(), match =>
            {
                var column = ParseColumn(match.Groups[4].Value);
                var row = long.Parse(match.Groups[6].Value, CultureInfo.InvariantCulture) - 1;
                var shiftedColumn = match.Groups[3].Value.Length > 0 ? column : column + columnOffset;
                var shiftedRow = match.Groups[5].Value.Length > 0 ? row : row + rowOffset;
                var prefix = match.Groups[1].Success ? $"'{match.Groups[1].Value}'!" : match.Groups[2].Success ? $"{match.Groups[2].Value}!" : string.Empty;
                if (shiftedColumn < 0 || shiftedColumn >= 16_384 || shiftedRow < 0 || shiftedRow >= 1_048_576) return $"{prefix}#REF!";
                return $"{prefix}{match.Groups[3].Value}{ColumnLabel((uint)shiftedColumn)}{match.Groups[5].Value}{shiftedRow + 1}";
            }));
            segment.Clear();
        }

        for (var index = 0; index < formula.Length; index++)
        {
            var character = formula[index];
            if (inString)
            {
                output.Append(character);
                if (character == '"')
                {
                    if (index + 1 < formula.Length && formula[index + 1] == '"') output.Append(formula[++index]);
                    else inString = false;
                }
                continue;
            }
            if (bracketDepth > 0)
            {
                output.Append(character);
                if (character == '[') bracketDepth++;
                else if (character == ']') bracketDepth--;
                continue;
            }
            if (character == '"')
            {
                Flush();
                output.Append(character);
                inString = true;
            }
            else if (character == '[')
            {
                Flush();
                output.Append(character);
                bracketDepth = 1;
            }
            else segment.Append(character);
        }
        Flush();
        return output.ToString();
    }

    private static RangeBounds ParseRange(string reference, string sheetName)
    {
        var pieces = reference.Split(':');
        if (pieces.Length is < 1 or > 2) throw new CodecException("invalid_cell_formula", $"Worksheet {sheetName} formula reference {reference} is not a bounded A1 range.", sheetName);
        var first = ParseCellReference(pieces[0], sheetName);
        var second = pieces.Length == 2 ? ParseCellReference(pieces[1], sheetName) : first;
        if (first.Row > second.Row || first.Column > second.Column)
            throw new CodecException("invalid_cell_formula", $"Worksheet {sheetName} formula reference {reference} must be top-left to bottom-right.", sheetName);
        return new RangeBounds(first.Row, first.Column, second.Row, second.Column);
    }

    private static (uint Row, uint Column) ParseCellReference(string? reference, string sheetName)
    {
        if (string.IsNullOrWhiteSpace(reference)) throw new CodecException("invalid_cell_formula", $"Worksheet {sheetName} has an empty formula cell reference.", sheetName);
        var match = Regex.Match(reference, @"^\$?([A-Za-z]{1,3})\$?([1-9]\d*)$", RegexOptions.CultureInvariant);
        if (!match.Success) throw new CodecException("invalid_cell_formula", $"Worksheet {sheetName} formula cell reference {reference} is invalid.", sheetName);
        var column = ParseColumn(match.Groups[1].Value);
        var row = uint.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture) - 1;
        if (column >= 16_384 || row >= 1_048_576) throw new CodecException("invalid_cell_formula", $"Worksheet {sheetName} formula cell reference {reference} exceeds XLSX limits.", sheetName);
        return (row, checked((uint)column));
    }

    private static long ParseColumn(string label)
    {
        var value = 0L;
        foreach (var character in label) value = checked(value * 26 + char.ToUpperInvariant(character) - 'A' + 1);
        return value - 1;
    }

    private static string ColumnLabel(uint column)
    {
        var number = checked((int)column + 1);
        Span<char> buffer = stackalloc char[3];
        var position = buffer.Length;
        while (number > 0)
        {
            number--;
            buffer[--position] = (char)('A' + number % 26);
            number /= 26;
        }
        return new string(buffer[position..]);
    }

    private static string CellReference(uint row, uint column) => $"{ColumnLabel(column)}{row + 1}";

    private static CodecException Invalid(Cell cell, string sheetName, string message) =>
        new("invalid_cell_formula", $"Cell {sheetName}!{cell.CellReference?.Value ?? "unknown"} {message}.", sheetName);

    private static CodecException Invalid(CellArtifact cell, string sheetName, string message) =>
        new("invalid_cell_formula", $"Cell {sheetName}!{CellReference(cell.Row, cell.Column)} {message}.", sheetName);

    private static CodecException Unsupported(Cell cell, string sheetName, string message) =>
        new("unsupported_cell_formula", $"Cell {sheetName}!{cell.CellReference?.Value ?? "unknown"} {message}.", sheetName);
}
