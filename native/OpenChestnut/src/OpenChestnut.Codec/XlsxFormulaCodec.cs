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
// A partial shared range emitted by Excel is the narrow exception: retain its
// original XML untouched, expose only the present cells' expanded formulas,
// and make every coordinate in that range source-bound/read-only.
internal sealed class XlsxFormulaCodec
{
    private const int MaxFormulaLength = 8_192;
    private const ulong MaxFormulaTopologyCells = 1_048_576;
    private static readonly Regex CellReferencePattern = new(
        @"(?<![A-Za-z0-9_.])(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_. ]*))!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_])",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private readonly string _sheetName;
    private readonly Dictionary<(uint Row, uint Column), FormulaRecord> _records;
    private readonly HashSet<(uint Row, uint Column)> _sourceBoundCoordinates;
    private readonly SourceBoundSharedFormulaRange[] _sourceBoundSharedFormulaRanges;
    private readonly XlsxDynamicArrayCodec _dynamicArrays;

    private sealed record FormulaRecord(string Formula, CellFormulaMetadata? Metadata, CellFormula Source);
    internal sealed record SourceBoundSharedFormulaRange(uint SharedIndex, string Reference)
    {
        internal string Display => $"si={SharedIndex} {Reference}";
    }
    private readonly record struct RangeBounds(uint Top, uint Left, uint Bottom, uint Right)
    {
        internal ulong CellCount => checked(((ulong)Bottom - Top + 1) * ((ulong)Right - Left + 1));
    }

    private XlsxFormulaCodec(
        string sheetName,
        Dictionary<(uint Row, uint Column), FormulaRecord> records,
        HashSet<(uint Row, uint Column)> sourceBoundCoordinates,
        IEnumerable<SourceBoundSharedFormulaRange> sourceBoundSharedFormulaRanges,
        XlsxDynamicArrayCodec dynamicArrays)
    {
        _sheetName = sheetName;
        _records = records;
        _sourceBoundCoordinates = sourceBoundCoordinates;
        _sourceBoundSharedFormulaRanges = sourceBoundSharedFormulaRanges.OrderBy(value => value.Reference, StringComparer.Ordinal).ThenBy(value => value.SharedIndex).ToArray();
        _dynamicArrays = dynamicArrays;
    }

    internal IReadOnlyList<SourceBoundSharedFormulaRange> SourceBoundSharedFormulaRanges => _sourceBoundSharedFormulaRanges;

    internal static XlsxFormulaCodec ForWorksheet(Worksheet worksheet, string sheetName, XlsxDynamicArrayCodec dynamicArrays)
    {
        var worksheetCells = worksheet.GetFirstChild<SheetData>()?.Descendants<Cell>().ToArray() ?? [];
        foreach (var cell in worksheetCells.Where(dynamicArrays.IsDynamic))
        {
            if (cell.CellFormula is null || FormulaType(cell.CellFormula) != CellFormulaValues.Array)
                throw Invalid(cell, sheetName, "has an XLDAPR dynamic-array marker but no array formula");
        }
        var formulas = worksheetCells
            .Where(cell => cell.CellFormula is not null)
            .Select(cell =>
            {
                var coordinate = ParseCellReference(cell.CellReference?.Value, sheetName);
                return (Coordinate: coordinate, Cell: cell, Formula: cell.CellFormula!);
            })
            .ToArray();
        var records = new Dictionary<(uint Row, uint Column), FormulaRecord>();
        var duplicate = formulas.GroupBy(item => item.Coordinate).FirstOrDefault(group => group.Count() > 1);
        if (duplicate is not null)
            throw new CodecException("invalid_cell_formula", $"Worksheet {sheetName} contains duplicate formula cell {CellReference(duplicate.Key.Row, duplicate.Key.Column)}.", sheetName);
        var formulasByCoordinate = formulas.ToDictionary(item => item.Coordinate);
        var topologyRanges = new List<(RangeBounds Bounds, string Owner, (uint Row, uint Column) Anchor, CellFormulaKind Kind)>();
        var sourceBoundCoordinates = new HashSet<(uint Row, uint Column)>();
        var sourceBoundSharedFormulaRanges = new List<SourceBoundSharedFormulaRange>();

        foreach (var entry in formulas.Where(item => FormulaType(item.Formula) != CellFormulaValues.Shared && FormulaType(item.Formula) != CellFormulaValues.Array))
        {
            var type = FormulaType(entry.Formula);
            if (type == CellFormulaValues.DataTable)
            {
                var metadata = ReadDataTableMetadata(entry.Formula, entry.Cell, entry.Coordinate, sheetName);
                var bounds = ParseRange(metadata.Reference, sheetName);
                records.Add(entry.Coordinate, new FormulaRecord(string.Empty, metadata, entry.Formula));
                topologyRanges.Add((bounds, $"data table {metadata.Reference}", entry.Coordinate, CellFormulaKind.DataTable));
                continue;
            }
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
            topologyRanges.Add((bounds, $"shared si={group.Key}", master.Coordinate, CellFormulaKind.Shared));
            if (master.Coordinate != (bounds.Top, bounds.Left))
                throw Invalid(master.Cell, sheetName, $"shared formula master must be the top-left cell of {reference}");
            var members = group.ToDictionary(item => item.Coordinate);
            if (members.Keys.Any(coordinate => coordinate.Row < bounds.Top || coordinate.Row > bounds.Bottom || coordinate.Column < bounds.Left || coordinate.Column > bounds.Right))
                throw new CodecException("invalid_cell_formula", $"Worksheet {sheetName} shared formula si={group.Key} has a member outside {reference}.", sheetName);
            var partial = (ulong)members.Count != bounds.CellCount;
            if (partial)
            {
                sourceBoundSharedFormulaRanges.Add(new SourceBoundSharedFormulaRange(group.Key, reference));
                for (var row = bounds.Top; row <= bounds.Bottom; row++)
                    for (var column = bounds.Left; column <= bounds.Right; column++)
                        sourceBoundCoordinates.Add((row, column));
            }
            var masterBody = ValidateFormulaBody(master.Formula.Text, master.Cell.CellReference?.Value ?? "cell", required: true);
            for (var row = bounds.Top; row <= bounds.Bottom; row++)
            {
                for (var column = bounds.Left; column <= bounds.Right; column++)
                {
                    if (!members.TryGetValue((row, column), out var member))
                    {
                        if (!partial)
                            throw new CodecException("invalid_cell_formula", $"Worksheet {sheetName} shared formula si={group.Key} is missing {CellReference(row, column)} from {reference}.", sheetName);
                        continue;
                    }
                    if (member.Coordinate != master.Coordinate && (!string.IsNullOrEmpty(member.Formula.Text) || member.Formula.Reference?.Value is { Length: > 0 }))
                        throw Invalid(member.Cell, sheetName, "shared formula follower must not contain formula text or ref");
                    var expanded = TranslateFormula(masterBody, master.Coordinate, member.Coordinate);
                    records.Add(member.Coordinate, new FormulaRecord(
                        $"={expanded}",
                        partial ? null : new CellFormulaMetadata
                        {
                            Kind = CellFormulaKind.Shared,
                            SharedIndex = group.Key,
                            Reference = reference,
                        },
                        member.Formula));
                }
            }
        }

        foreach (var entry in formulas.Where(item => FormulaType(item.Formula) == CellFormulaValues.Array))
        {
            if (entry.Formula.Reference?.Value is not { Length: > 0 } reference)
                throw Invalid(entry.Cell, sheetName, dynamicArrays.IsDynamic(entry.Cell) ? "dynamic array formula requires ref" : "legacy array formula requires ref");
            var bounds = ParseRange(reference, sheetName);
            var kind = dynamicArrays.IsDynamic(entry.Cell) ? CellFormulaKind.DynamicArray : CellFormulaKind.Array;
            var label = kind == CellFormulaKind.DynamicArray ? "dynamic array" : "legacy array";
            topologyRanges.Add((bounds, $"{label} {CellReference(entry.Coordinate.Row, entry.Coordinate.Column)}", entry.Coordinate, kind));
            if (entry.Coordinate != (bounds.Top, bounds.Left))
                throw Invalid(entry.Cell, sheetName, $"{label} formula must be anchored at the top-left cell of {reference}");
            var body = ValidateFormulaBody(entry.Formula.Text, entry.Cell.CellReference?.Value ?? "cell", required: true);
            records.Add(entry.Coordinate, new FormulaRecord($"={body}", new CellFormulaMetadata
            {
                Kind = kind,
                Reference = reference,
            }, entry.Formula));
        }

        var occupied = new Dictionary<(uint Row, uint Column), string>();
        ulong topologyCellCount = 0;
        foreach (var topology in topologyRanges)
        {
            topologyCellCount = checked(topologyCellCount + topology.Bounds.CellCount);
            if (topologyCellCount > MaxFormulaTopologyCells)
                throw new CodecException("invalid_cell_formula", $"Worksheet {sheetName} native formula topology exceeds {MaxFormulaTopologyCells} cells.", sheetName);
            for (var row = topology.Bounds.Top; row <= topology.Bounds.Bottom; row++)
            {
                for (var column = topology.Bounds.Left; column <= topology.Bounds.Right; column++)
                {
                    var coordinate = (row, column);
                    if (occupied.TryGetValue(coordinate, out var previous) && previous != topology.Owner)
                        throw new CodecException("invalid_cell_formula", $"Worksheet {sheetName} formula range owned by {topology.Owner} overlaps {previous} at {CellReference(row, column)}.", sheetName);
                    occupied[coordinate] = topology.Owner;
                    if ((topology.Kind is CellFormulaKind.Array or CellFormulaKind.DynamicArray or CellFormulaKind.DataTable) && coordinate != topology.Anchor && formulasByCoordinate.TryGetValue(coordinate, out var nested))
                        throw Invalid(nested.Cell, sheetName, $"must not contain another formula inside range owned by {topology.Owner}");
                }
            }
        }

        return new XlsxFormulaCodec(sheetName, records, sourceBoundCoordinates, sourceBoundSharedFormulaRanges, dynamicArrays);
    }

    internal static void ValidateArtifact(WorksheetArtifact sheet)
    {
        var coordinates = new HashSet<(uint Row, uint Column)>();
        foreach (var cell in sheet.Cells)
        {
            if (!coordinates.Add((cell.Row, cell.Column)))
                throw new CodecException("duplicate_cell", $"Worksheet {sheet.Name} contains duplicate cell {CellReference(cell.Row, cell.Column)}.", sheet.Name);
            ValidateFormulaBody(cell.Formula, $"{sheet.Name}!{CellReference(cell.Row, cell.Column)}", required: cell.FormulaMetadata is not null && cell.FormulaMetadata.Kind != CellFormulaKind.DataTable);
            if (cell.FormulaMetadata is null) continue;
            if (cell.FormulaMetadata.Kind == CellFormulaKind.Unspecified)
                throw Invalid(cell, sheet.Name, "formula_metadata.kind must be shared, array, dynamic_array, or data_table");
            if (cell.FormulaMetadata.Kind is not (CellFormulaKind.Shared or CellFormulaKind.Array or CellFormulaKind.DynamicArray or CellFormulaKind.DataTable))
                throw Invalid(cell, sheet.Name, "formula_metadata.kind is unsupported");
            if (string.IsNullOrWhiteSpace(cell.FormulaMetadata.Reference))
                throw Invalid(cell, sheet.Name, "formula metadata requires reference");
            if (cell.FormulaMetadata.Kind == CellFormulaKind.DataTable) ValidateDataTableMetadata(cell, sheet.Name);
        }
        var cellsByCoordinate = sheet.Cells.ToDictionary(cell => (cell.Row, cell.Column));

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
        var topologyRoots = sheet.Cells
            .Where(cell => cell.FormulaMetadata?.Kind == CellFormulaKind.Shared)
            .GroupBy(cell => cell.FormulaMetadata!.SharedIndex)
            .Select(group => group.First())
            .Concat(sheet.Cells.Where(cell => cell.FormulaMetadata?.Kind is CellFormulaKind.Array or CellFormulaKind.DynamicArray or CellFormulaKind.DataTable));
        ulong topologyCellCount = 0;
        foreach (var cell in topologyRoots)
        {
            var metadata = cell.FormulaMetadata!;
            var bounds = ParseRange(metadata.Reference, sheet.Name);
            topologyCellCount = checked(topologyCellCount + bounds.CellCount);
            if (topologyCellCount > MaxFormulaTopologyCells)
                throw Invalid(cell, sheet.Name, $"native formula topology exceeds {MaxFormulaTopologyCells} cells");
            var arrayKind = metadata.Kind is CellFormulaKind.Array or CellFormulaKind.DynamicArray;
            var dataTable = metadata.Kind == CellFormulaKind.DataTable;
            var label = metadata.Kind == CellFormulaKind.DynamicArray ? "dynamic array" : "legacy array";
            var owner = metadata.Kind == CellFormulaKind.Shared ? $"shared si={metadata.SharedIndex}" : dataTable ? $"data table {metadata.Reference}" : $"{label} {CellReference(cell.Row, cell.Column)}";
            if ((arrayKind || dataTable) && (cell.Row != bounds.Top || cell.Column != bounds.Left))
                throw Invalid(cell, sheet.Name, $"{(dataTable ? "data table" : label)} formula must be anchored at the top-left cell of {metadata.Reference}");
            if (arrayKind && metadata.SharedIndex != 0)
                throw Invalid(cell, sheet.Name, $"{label} formula must not set shared_index");
            for (var row = bounds.Top; row <= bounds.Bottom; row++)
            {
                for (var column = bounds.Left; column <= bounds.Right; column++)
                {
                    if (topology.TryGetValue((row, column), out var previous) && previous != owner)
                        throw Invalid(cell, sheet.Name, $"formula range {metadata.Reference} overlaps {previous}");
                    topology[(row, column)] = owner;
                    if ((arrayKind || dataTable) && (row != cell.Row || column != cell.Column) &&
                        cellsByCoordinate.TryGetValue((row, column), out var nested) && !string.IsNullOrWhiteSpace(nested.Formula))
                        throw Invalid(nested, sheet.Name, $"must not contain another formula inside {(dataTable ? "data table" : label)} range {metadata.Reference}");
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
        if (string.IsNullOrWhiteSpace(source.Formula) && source.FormulaMetadata?.Kind != CellFormulaKind.DataTable) return null;
        var body = ValidateFormulaBody(source.Formula, CellReference(source.Row, source.Column), required: source.FormulaMetadata?.Kind != CellFormulaKind.DataTable);
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
            CellFormulaKind.DynamicArray => new CellFormula(body)
            {
                FormulaType = CellFormulaValues.Array,
                Reference = metadata.Reference,
            },
            CellFormulaKind.DataTable when metadata.Editable => new CellFormula
            {
                FormulaType = CellFormulaValues.DataTable,
                Reference = metadata.Reference,
                DataTable2D = metadata.TwoVariable,
                DataTableRow = metadata.TwoVariable ? null : metadata.RowOriented,
                R1 = string.IsNullOrEmpty(metadata.RowInput) ? null : metadata.RowInput,
                R2 = string.IsNullOrEmpty(metadata.ColumnInput) ? null : metadata.ColumnInput,
            },
            CellFormulaKind.DataTable => throw new CodecException("unsupported_cell_formula_edit", $"Cell worksheet!{CellReference(source.Row, source.Column)} data table is source-bound and read-only."),
            _ => throw Invalid(source, "worksheet", "formula_metadata.kind must be shared, array, dynamic_array, or data_table"),
        };
    }

    internal void Apply(Cell target, CellArtifact desired)
    {
        AssertCellEditable(desired);
        var coordinate = (desired.Row, desired.Column);
        if (_records.TryGetValue(coordinate, out var current) && HasUnmodeledAttributes(current.Source))
            throw new CodecException("unsupported_cell_formula_edit", $"Cell {_sheetName}!{CellReference(desired.Row, desired.Column)} has unmodeled formula attributes and cannot be edited losslessly.", _sheetName);
        _dynamicArrays.ApplyFormulaMetadata(target, desired, sourceBound: true);
        target.CellFormula = Build(desired);
    }

    internal void AssertCellEditable(CellArtifact desired)
    {
        if (!_sourceBoundCoordinates.Contains((desired.Row, desired.Column))) return;
        throw new CodecException(
            "unsupported_cell_formula_edit",
            $"Cell {_sheetName}!{CellReference(desired.Row, desired.Column)} belongs to a partial native shared-formula range and is source-bound/read-only.",
            _sheetName);
    }

    internal static bool SemanticallyEqual(CellArtifact left, CellArtifact right)
    {
        if (!string.Equals(NormalizeFormula(left.Formula), NormalizeFormula(right.Formula), StringComparison.Ordinal)) return false;
        var leftMetadata = left.FormulaMetadata;
        var rightMetadata = right.FormulaMetadata;
        if (leftMetadata is null || rightMetadata is null) return leftMetadata is null && rightMetadata is null;
        return leftMetadata.Kind == rightMetadata.Kind && leftMetadata.SharedIndex == rightMetadata.SharedIndex &&
               string.Equals(leftMetadata.Reference, rightMetadata.Reference, StringComparison.OrdinalIgnoreCase) &&
               string.Equals(leftMetadata.RowInput, rightMetadata.RowInput, StringComparison.OrdinalIgnoreCase) &&
               string.Equals(leftMetadata.ColumnInput, rightMetadata.ColumnInput, StringComparison.OrdinalIgnoreCase) &&
               leftMetadata.RowOriented == rightMetadata.RowOriented && leftMetadata.TwoVariable == rightMetadata.TwoVariable &&
               leftMetadata.Editable == rightMetadata.Editable;
    }

    private static bool HasUnmodeledAttributes(CellFormula formula)
    {
        var type = FormulaType(formula);
        var allowed = type == CellFormulaValues.Shared
            ? new HashSet<string>(["t", "si", "ref"], StringComparer.Ordinal)
            : type == CellFormulaValues.Array
                ? new HashSet<string>(["t", "ref"], StringComparer.Ordinal)
                : type == CellFormulaValues.DataTable
                    ? new HashSet<string>(["t", "ref", "dt2D", "dtr", "r1", "r2"], StringComparer.Ordinal)
                : new HashSet<string>(["t"], StringComparer.Ordinal);
        return formula.GetAttributes().Any(attribute => !allowed.Contains(attribute.LocalName));
    }

    private static CellFormulaMetadata ReadDataTableMetadata(CellFormula formula, Cell cell, (uint Row, uint Column) coordinate, string sheetName)
    {
        if (formula.Reference?.Value is not { Length: > 0 } reference)
            throw Invalid(cell, sheetName, "data table requires ref");
        var bounds = ParseRange(reference, sheetName);
        if (coordinate != (bounds.Top, bounds.Left))
            throw Invalid(cell, sheetName, $"data table must be anchored at the top-left cell of {reference}");
        var twoVariable = formula.DataTable2D?.Value == true;
        var rowOriented = formula.DataTableRow?.Value == true;
        var rowInput = NormalizeDataTableInput(formula.R1?.Value, sheetName);
        var columnInput = NormalizeDataTableInput(formula.R2?.Value, sheetName);
        ValidateDataTableInputs(cell, sheetName, rowInput, columnInput, rowOriented, twoVariable);
        return new CellFormulaMetadata
        {
            Kind = CellFormulaKind.DataTable,
            Reference = reference,
            RowInput = rowInput,
            ColumnInput = columnInput,
            RowOriented = rowOriented,
            TwoVariable = twoVariable,
            Editable = string.IsNullOrWhiteSpace(formula.Text) && !HasUnmodeledAttributes(formula) &&
                       formula.Input1Deleted?.Value != true && formula.Input2Deleted?.Value != true,
        };
    }

    private static void ValidateDataTableMetadata(CellArtifact cell, string sheetName)
    {
        var metadata = cell.FormulaMetadata!;
        if (!string.IsNullOrWhiteSpace(cell.Formula)) throw Invalid(cell, sheetName, "data table formula body must be empty");
        var bounds = ParseRange(metadata.Reference, sheetName);
        if (cell.Row != bounds.Top || cell.Column != bounds.Left)
            throw Invalid(cell, sheetName, $"data table must be anchored at the top-left cell of {metadata.Reference}");
        var rowInput = NormalizeDataTableInput(metadata.RowInput, sheetName);
        var columnInput = NormalizeDataTableInput(metadata.ColumnInput, sheetName);
        if (!string.Equals(rowInput, metadata.RowInput, StringComparison.Ordinal) || !string.Equals(columnInput, metadata.ColumnInput, StringComparison.Ordinal))
            throw Invalid(cell, sheetName, "data table input addresses must be normalized local A1 cells");
        ValidateDataTableInputs(cell, sheetName, rowInput, columnInput, metadata.RowOriented, metadata.TwoVariable);
    }

    private static void ValidateDataTableInputs(Cell cell, string sheetName, string rowInput, string columnInput, bool rowOriented, bool twoVariable)
    {
        if (twoVariable)
        {
            if (rowInput.Length == 0 || columnInput.Length == 0 || rowOriented)
                throw Invalid(cell, sheetName, "two-variable data table requires row and column inputs and must not set dtr");
            return;
        }
        if (rowOriented)
        {
            if (rowInput.Length == 0 || columnInput.Length != 0)
                throw Invalid(cell, sheetName, "row-oriented data table requires r1 only");
            return;
        }
        if (rowInput.Length != 0 || columnInput.Length == 0)
            throw Invalid(cell, sheetName, "column-oriented data table requires r2 only");
    }

    private static void ValidateDataTableInputs(CellArtifact cell, string sheetName, string rowInput, string columnInput, bool rowOriented, bool twoVariable)
    {
        if (twoVariable)
        {
            if (rowInput.Length == 0 || columnInput.Length == 0 || rowOriented)
                throw Invalid(cell, sheetName, "two-variable data table requires row and column inputs and must not set row_oriented");
            return;
        }
        if (rowOriented)
        {
            if (rowInput.Length == 0 || columnInput.Length != 0)
                throw Invalid(cell, sheetName, "row-oriented data table requires row_input only");
            return;
        }
        if (rowInput.Length != 0 || columnInput.Length == 0)
            throw Invalid(cell, sheetName, "column-oriented data table requires column_input only");
    }

    private static string NormalizeDataTableInput(string? value, string sheetName)
    {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        var coordinate = ParseCellReference(value, sheetName);
        return CellReference(coordinate.Row, coordinate.Column);
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
            var sourceText = segment.ToString();
            output.Append(CellReferencePattern.Replace(sourceText, match =>
            {
                if (sourceText.AsSpan(match.Index + match.Length).TrimStart().StartsWith("(")) return match.Value;
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
