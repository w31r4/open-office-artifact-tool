using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Pivot caches are derived package projections. The public semantic contract
// stays on SpreadsheetPivotTableArtifact; this codec alone owns the workbook,
// worksheet, PivotTable, cache-definition, and cache-record relationship graph.
internal sealed class XlsxPivotTableCodec
{
    private const int MaxPivots = 16_384;
    private const int MaxRowFields = 8;
    private const int MaxValueFields = 32;
    private const int MaxFilterItems = 1_024;
    private const int DataLayoutFieldIndex = -2;
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly Regex A1Range = new(
        "^\\$?(?<firstColumn>[A-Za-z]{1,3})\\$?(?<firstRow>[1-9][0-9]*)(?::\\$?(?<lastColumn>[A-Za-z]{1,3})\\$?(?<lastRow>[1-9][0-9]*))?$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private readonly WorkbookPart _workbookPart;

    internal XlsxPivotTableCodec(WorkbookPart workbookPart) => _workbookPart = workbookPart;

    internal void Apply(IReadOnlyList<(WorksheetPart Part, WorksheetArtifact Artifact)> worksheets, bool sourceBound)
    {
        var count = worksheets.Sum(item => item.Artifact.PivotTables.Count);
        if (count > MaxPivots) throw Invalid($"Workbook has {count} PivotTables and exceeds the {MaxPivots} PivotTable budget.");
        if (!sourceBound)
        {
            if (count == 0) return;
            var duplicateName = worksheets.SelectMany(item => item.Artifact.PivotTables)
                .GroupBy(pivot => pivot.Name, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault(group => group.Count() > 1)?.Key;
            if (duplicateName is not null) throw Invalid($"PivotTable name {duplicateName} must be unique across the workbook.");
            var byId = worksheets.ToDictionary(item => item.Artifact.Id, StringComparer.Ordinal);
            uint cacheId = 1;
            foreach (var target in worksheets)
                foreach (var pivot in target.Artifact.PivotTables)
                    Author(target, pivot, byId, cacheId++);
            _workbookPart.Workbook!.Save();
            return;
        }

        foreach (var worksheet in worksheets)
        {
            var records = Scan(worksheet.Part, worksheet.Artifact.Id);
            if (records.Count != worksheet.Artifact.PivotTables.Count)
                throw new CodecException("invalid_spreadsheet_pivot_topology", $"Worksheet {worksheet.Artifact.Name} source-bound recognized PivotTable count cannot change from {records.Count} to {worksheet.Artifact.PivotTables.Count}.", PartPath(worksheet.Part));
            for (var index = 0; index < records.Count; index++)
            {
                var target = worksheet.Artifact.PivotTables[index];
                var record = records[index];
                var binding = target.Source ?? throw Invalid($"Imported PivotTable {target.Name} is missing its source binding.");
                if (!string.Equals(target.Id, record.Artifact.Id, StringComparison.Ordinal) ||
                    binding.PivotOrdinal != record.Ordinal ||
                    binding.CacheId != record.CacheId ||
                    !string.Equals(binding.WorksheetPartPath, PartPath(worksheet.Part), StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(binding.PivotTablePartPath, PartPath(record.PivotPart), StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(binding.PivotTableXmlSha256, Hash(record.PivotPart.PivotTableDefinition!.OuterXml), StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(binding.CacheDefinitionPartPath, PartPath(record.CachePart), StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(binding.CacheDefinitionXmlSha256, Hash(record.CachePart.PivotCacheDefinition!.OuterXml), StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(binding.CacheRecordsPartPath, record.RecordsPart is null ? "" : PartPath(record.RecordsPart), StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(binding.CacheRecordsXmlSha256, record.RecordsPart?.PivotCacheRecords is null ? "" : Hash(record.RecordsPart.PivotCacheRecords.OuterXml), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("invalid_spreadsheet_pivot_source", $"PivotTable {target.Name} no longer matches its validated package locator.", PartPath(record.PivotPart));
                if (!string.Equals(binding.SemanticSha256, SemanticHash(target), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("unsupported_spreadsheet_pivot_edit", $"Imported PivotTable {target.Name} is read-only in the first native profile.", PartPath(record.PivotPart));
            }
        }
    }

    internal IReadOnlyList<SpreadsheetPivotTableArtifact> Read(WorksheetPart worksheetPart, string worksheetId) =>
        Scan(worksheetPart, worksheetId).Select(record => record.Artifact).ToArray();

    private void Author(
        (WorksheetPart Part, WorksheetArtifact Artifact) target,
        SpreadsheetPivotTableArtifact pivot,
        IReadOnlyDictionary<string, (WorksheetPart Part, WorksheetArtifact Artifact)> worksheets,
        uint cacheId)
    {
        if (pivot.Source is not null) throw Invalid($"Source-free PivotTable {pivot.Name} cannot carry a source binding.");
        ValidateSemantic(pivot);
        if (!worksheets.TryGetValue(pivot.SourceWorksheetId, out var source)) throw Invalid($"PivotTable {pivot.Name} references missing source worksheet {pivot.SourceWorksheetId}.");
        var sourceRange = ParseRange(pivot.SourceReference, $"PivotTable {pivot.Name} source_reference");
        var targetRange = ParseRange(pivot.TargetReference, $"PivotTable {pivot.Name} target_reference");
        if (sourceRange.RowCount < 2 || sourceRange.ColumnCount < 2) throw Invalid($"PivotTable {pivot.Name} source range must include headers and data.");

        var sourceCells = source.Artifact.Cells.ToDictionary(cell => (cell.Row, cell.Column));
        var headers = Enumerable.Range(0, sourceRange.ColumnCount)
            .Select(offset => CellText(sourceCells.GetValueOrDefault((sourceRange.Top, checked((uint)(sourceRange.Left + offset))))).Trim())
            .ToArray();
        if (headers.Any(string.IsNullOrWhiteSpace) || headers.Distinct(StringComparer.Ordinal).Count() != headers.Length)
            throw Invalid($"PivotTable {pivot.Name} source headers must be non-empty and unique.");
        var rowIndexes = pivot.RowFields.Select(field => Array.IndexOf(headers, field)).ToArray();
        var columnIndex = pivot.ColumnFields.Count == 0 ? -1 : Array.IndexOf(headers, pivot.ColumnFields[0]);
        var valueIndexes = pivot.ValueFields.Select(value => Array.IndexOf(headers, value.Field)).ToArray();
        if (rowIndexes.Any(index => index < 0) || (pivot.ColumnFields.Count > 0 && columnIndex < 0) || valueIndexes.Any(index => index < 0))
            throw Invalid($"PivotTable {pivot.Name} contains a field outside its source headers.");

        var rows = new List<CacheValue[]>();
        for (var row = sourceRange.Top + 1; row <= sourceRange.Bottom; row++)
        {
            var values = new CacheValue[sourceRange.ColumnCount];
            for (var offset = 0; offset < sourceRange.ColumnCount; offset++)
                values[offset] = CacheValue.FromCell(sourceCells.GetValueOrDefault((row, checked((uint)(sourceRange.Left + offset)))));
            rows.Add(values);
        }
        var shared = Enumerable.Range(0, sourceRange.ColumnCount)
            .Select(column => DistinctValues(rows.Select(row => row[column])))
            .ToArray();

        var filters = ResolveItemFilters(pivot, headers, shared);
        var activeRows = rows.Where(row => filters.All(filter => filter.Value.Visible(row[filter.Key]))).ToArray();
        if (activeRows.Length == 0) throw new CodecException("unsupported_spreadsheet_pivot_filter", $"PivotTable {pivot.Name} item filters hide every source row.");
        var rowItemIndexes = ActiveItemTuples(activeRows, rowIndexes, shared);
        var columnItemIndexes = columnIndex >= 0 ? ActiveItemIndexes(activeRows, columnIndex, shared[columnIndex]) : [];

        var expectedRows = 1 + rowItemIndexes.Count + (pivot.ColumnGrandTotals ? 1 : 0);
        var expectedColumns = rowIndexes.Length + (columnIndex >= 0
            ? (columnItemIndexes.Count + (pivot.RowGrandTotals ? 1 : 0)) * valueIndexes.Length
            : valueIndexes.Length);
        if (targetRange.RowCount != expectedRows || targetRange.ColumnCount != expectedColumns)
            throw Invalid($"PivotTable {pivot.Name} target range must be {expectedRows}x{expectedColumns} for its cached axis members and grand-total policy.");
        var targetCells = target.Artifact.Cells.ToDictionary(cell => (cell.Row, cell.Column));
        for (var row = targetRange.Top; row <= targetRange.Bottom; row++)
            for (var column = targetRange.Left; column <= targetRange.Right; column++)
                if (!targetCells.ContainsKey((row, column))) throw Invalid($"PivotTable {pivot.Name} cached output is missing cell {ColumnLabel(column)}{row + 1}.");

        var cachePart = _workbookPart.AddNewPart<PivotTableCacheDefinitionPart>();
        PivotTableCacheRecordsPart? recordsPart = null;
        string? recordsRelationshipId = null;
        if (pivot.RefreshPolicy?.SaveData != false)
        {
            recordsPart = cachePart.AddNewPart<PivotTableCacheRecordsPart>();
            recordsRelationshipId = cachePart.GetIdOfPart(recordsPart);
            recordsPart.PivotCacheRecords = new PivotCacheRecords(BuildCacheRecords(rows, shared));
        }
        cachePart.PivotCacheDefinition = new PivotCacheDefinition(BuildCacheDefinition(pivot, source.Artifact.Name, headers, shared, rows.Count, recordsRelationshipId));

        var pivotPart = target.Part.AddNewPart<PivotTablePart>();
        pivotPart.AddPart(cachePart);
        pivotPart.PivotTableDefinition = new PivotTableDefinition(BuildPivotTableDefinition(
            pivot, headers, shared, rowIndexes, columnIndex, valueIndexes, filters,
            rowItemIndexes, columnItemIndexes, cacheId, targetRange));

        var pivotCaches = _workbookPart.Workbook!.GetFirstChild<PivotCaches>();
        if (pivotCaches is null)
        {
            pivotCaches = new PivotCaches();
            var sheets = _workbookPart.Workbook.GetFirstChild<Sheets>();
            if (sheets is null) _workbookPart.Workbook.Append(pivotCaches);
            else _workbookPart.Workbook.InsertAfter(pivotCaches, sheets);
        }
        pivotCaches.Append(new PivotCache { CacheId = cacheId, Id = _workbookPart.GetIdOfPart(cachePart) });
    }

    private IReadOnlyList<PivotRecord> Scan(WorksheetPart worksheetPart, string worksheetId)
    {
        var output = new List<PivotRecord>();
        var parts = worksheetPart.PivotTableParts.ToArray();
        for (var ordinal = 0; ordinal < parts.Length; ordinal++)
        {
            var pivotPart = parts[ordinal];
            var cachePart = pivotPart.PivotTableCacheDefinitionPart;
            if (pivotPart.PivotTableDefinition is null || cachePart?.PivotCacheDefinition is null) continue;
            if (!TryRead(pivotPart, cachePart, worksheetId, checked((uint)ordinal), out var artifact, out var cacheId)) continue;
            var recordsPart = cachePart.PivotTableCacheRecordsPart;
            artifact.Source = new SpreadsheetPivotTableSourceBinding
            {
                WorksheetPartPath = PartPath(worksheetPart),
                PivotTablePartPath = PartPath(pivotPart),
                PivotTableXmlSha256 = Hash(pivotPart.PivotTableDefinition.OuterXml),
                CacheDefinitionPartPath = PartPath(cachePart),
                CacheDefinitionXmlSha256 = Hash(cachePart.PivotCacheDefinition.OuterXml),
                CacheRecordsPartPath = recordsPart is null ? "" : PartPath(recordsPart),
                CacheRecordsXmlSha256 = recordsPart?.PivotCacheRecords is null ? "" : Hash(recordsPart.PivotCacheRecords.OuterXml),
                PivotOrdinal = checked((uint)ordinal),
                CacheId = cacheId,
                Editable = false,
            };
            artifact.Source.SemanticSha256 = SemanticHash(artifact);
            output.Add(new PivotRecord(artifact, pivotPart, cachePart, recordsPart, checked((uint)ordinal), cacheId));
        }
        return output;
    }

    private bool TryRead(
        PivotTablePart pivotPart,
        PivotTableCacheDefinitionPart cachePart,
        string worksheetId,
        uint ordinal,
        out SpreadsheetPivotTableArtifact artifact,
        out uint cacheId)
    {
        artifact = new SpreadsheetPivotTableArtifact();
        cacheId = 0;
        XDocument pivotDocument;
        XDocument cacheDocument;
        try
        {
            pivotDocument = XDocument.Parse(pivotPart.PivotTableDefinition!.OuterXml, LoadOptions.PreserveWhitespace);
            cacheDocument = XDocument.Parse(cachePart.PivotCacheDefinition!.OuterXml, LoadOptions.PreserveWhitespace);
        }
        catch
        {
            return false;
        }
        var root = pivotDocument.Root;
        var cacheRoot = cacheDocument.Root;
        if (root?.Name != Main + "pivotTableDefinition" || cacheRoot?.Name != Main + "pivotCacheDefinition") return false;
        if (!uint.TryParse(root.Attribute("cacheId")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out cacheId)) return false;
        var parsedCacheId = cacheId;
        var cacheEntries = _workbookPart.Workbook?.PivotCaches?.Elements<PivotCache>().Where(item => item.CacheId?.Value == parsedCacheId).ToArray() ?? [];
        if (cacheEntries.Length != 1 || cacheEntries[0].Id?.Value is not { Length: > 0 } cacheRelationshipId ||
            !ReferenceEquals(_workbookPart.GetPartById(cacheRelationshipId), cachePart)) return false;
        var recordsPart = cachePart.PivotTableCacheRecordsPart;
        var recordsRelationshipId = cacheRoot.Attribute(Relationships + "id")?.Value;
        if (recordsPart is null ? !string.IsNullOrEmpty(recordsRelationshipId) :
            string.IsNullOrEmpty(recordsRelationshipId) || !ReferenceEquals(cachePart.GetPartById(recordsRelationshipId), recordsPart)) return false;
        var location = root.Elements(Main + "location").SingleOrDefault();
        var pivotFields = root.Elements(Main + "pivotFields").SingleOrDefault()?.Elements(Main + "pivotField").ToArray();
        var rowFieldElements = root.Elements(Main + "rowFields").SingleOrDefault()?.Elements(Main + "field").ToArray() ?? [];
        var columnFieldElements = root.Elements(Main + "colFields").SingleOrDefault()?.Elements(Main + "field").ToArray() ?? [];
        var columnItems = root.Elements(Main + "colItems").SingleOrDefault()?.Elements(Main + "i").ToArray();
        var dataFields = root.Elements(Main + "dataFields").SingleOrDefault()?.Elements(Main + "dataField").ToArray() ?? [];
        var cacheFields = cacheRoot.Elements(Main + "cacheFields").SingleOrDefault()?.Elements(Main + "cacheField").ToArray();
        var worksheetSource = cacheRoot.Elements(Main + "cacheSource").SingleOrDefault()?.Elements(Main + "worksheetSource").SingleOrDefault();
        if (location is null || pivotFields is null || ReadBoolean(root.Attribute("dataOnRows"), false) ||
            rowFieldElements.Length is < 1 or > MaxRowFields || dataFields.Length is < 1 or > MaxValueFields ||
            cacheFields is null || worksheetSource is null || pivotFields.Length != cacheFields.Length) return false;
        // rowItems/colItems are optional materialized axis caches. LibreOffice
        // 24.2 intentionally omits both while retaining the canonical x=-2
        // data-layout field and ordered dataFields. Validate a list when the
        // host writes one, but do not require it to recover the semantic axes.
        if (dataFields.Length > 1 && columnItems is not null && columnItems.Length < dataFields.Length) return false;
        var rowIndexes = new List<int>();
        foreach (var rowField in rowFieldElements)
        {
            if (!int.TryParse(rowField.Attribute("x")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var rowIndex) ||
                rowIndex < 0 || rowIndex >= cacheFields.Length || rowIndexes.Contains(rowIndex)) return false;
            rowIndexes.Add(rowIndex);
        }
        foreach (var rowIndex in rowIndexes)
        {
            if (pivotFields[rowIndex].Attribute("axis")?.Value != "axisRow") return false;
            if (rowIndexes.Count > 1)
            {
                if (!TryReadOptionalBoolean(pivotFields[rowIndex].Attribute("compact"), out var compact) || compact != false ||
                    !TryReadOptionalBoolean(pivotFields[rowIndex].Attribute("outline"), out var outline) || outline == true ||
                    !TryReadOptionalBoolean(pivotFields[rowIndex].Attribute("defaultSubtotal"), out var defaultSubtotal) || defaultSubtotal != false) return false;
            }
        }

        var columnIndex = -1;
        var dataLayoutFields = 0;
        foreach (var columnField in columnFieldElements)
        {
            if (!int.TryParse(columnField.Attribute("x")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedColumnIndex)) return false;
            if (parsedColumnIndex == DataLayoutFieldIndex)
            {
                dataLayoutFields++;
                continue;
            }
            if (parsedColumnIndex < 0 || parsedColumnIndex >= cacheFields.Length || columnIndex >= 0) return false;
            columnIndex = parsedColumnIndex;
        }
        if (columnIndex >= 0 && (rowIndexes.Contains(columnIndex) || pivotFields[columnIndex].Attribute("axis")?.Value != "axisCol")) return false;
        if (dataFields.Length == 1 ? dataLayoutFields != 0 : dataLayoutFields != 1) return false;

        var parsedValueFields = new List<(int Index, string Name, SpreadsheetPivotAggregation Aggregation)>();
        foreach (var dataField in dataFields)
        {
            if (!int.TryParse(dataField.Attribute("fld")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var valueIndex) || valueIndex < 0 || valueIndex >= cacheFields.Length) return false;
            if (!TryAggregation(dataField.Attribute("subtotal")?.Value, out var aggregation)) return false;
            if (!ReadBoolean(pivotFields[valueIndex].Attribute("dataField"), false)) return false;
            parsedValueFields.Add((valueIndex, dataField.Attribute("name")?.Value ?? "", aggregation));
        }
        var headers = cacheFields.Select(field => field.Attribute("name")?.Value ?? "").ToArray();
        if (headers.Any(string.IsNullOrWhiteSpace) || headers.Distinct(StringComparer.Ordinal).Count() != headers.Length) return false;
        var itemFilters = new List<SpreadsheetPivotItemFilterArtifact>();
        foreach (var fieldIndex in columnIndex >= 0 ? rowIndexes.Append(columnIndex) : rowIndexes)
        {
            if (!TryReadItemFilter(pivotFields[fieldIndex], cacheFields[fieldIndex], headers[fieldIndex], out var filter)) return false;
            if (filter is not null) itemFilters.Add(filter);
        }
        var sourceSheetName = worksheetSource.Attribute("sheet")?.Value;
        var sourceWorksheet = _workbookPart.Workbook?.Sheets?.Elements<Sheet>().SingleOrDefault(sheet => string.Equals(sheet.Name?.Value, sourceSheetName, StringComparison.Ordinal));
        if (sourceWorksheet?.Id?.Value is not { Length: > 0 } sourceRelationshipId || _workbookPart.GetPartById(sourceRelationshipId) is not WorksheetPart sourcePart) return false;
        var sourceWorksheetId = $"worksheet/{Array.IndexOf(_workbookPart.Workbook!.Sheets!.Elements<Sheet>().ToArray(), sourceWorksheet) + 1}";

        artifact = new SpreadsheetPivotTableArtifact
        {
            Id = $"{worksheetId}/pivot/{ordinal + 1}",
            Name = root.Attribute("name")?.Value ?? $"PivotTable{ordinal + 1}",
            SourceWorksheetId = sourceWorksheetId,
            SourceReference = worksheetSource.Attribute("ref")?.Value ?? "",
            TargetReference = location.Attribute("ref")?.Value ?? "",
            RowGrandTotals = ReadBoolean(root.Attribute("rowGrandTotals"), true),
            ColumnGrandTotals = ReadBoolean(root.Attribute("colGrandTotals"), true),
            RefreshPolicy = new SpreadsheetPivotRefreshPolicyArtifact
            {
                RefreshOnLoad = ReadBoolean(cacheRoot.Attribute("refreshOnLoad"), false),
                SaveData = ReadBoolean(cacheRoot.Attribute("saveData"), true),
                EnableRefresh = ReadBoolean(cacheRoot.Attribute("enableRefresh"), true),
                Invalid = ReadBoolean(cacheRoot.Attribute("invalid"), false),
                MissingItemsLimit = ReadUInt(cacheRoot.Attribute("missingItemsLimit")),
                RefreshedBy = cacheRoot.Attribute("refreshedBy")?.Value ?? "",
                RefreshedDateIso = cacheRoot.Attribute("refreshedDateIso")?.Value ?? "",
            },
        };
        artifact.RowFields.Add(rowIndexes.Select(index => headers[index]));
        if (columnIndex >= 0) artifact.ColumnFields.Add(headers[columnIndex]);
        artifact.ValueFields.Add(parsedValueFields.Select(value => new SpreadsheetPivotValueFieldArtifact
        {
            Field = headers[value.Index],
            Name = value.Name,
            Aggregation = value.Aggregation,
        }));
        artifact.ItemFilters.Add(itemFilters);
        try
        {
            ValidateSemantic(artifact);
            ParseRange(artifact.SourceReference, $"PivotTable {artifact.Name} source_reference");
            ParseRange(artifact.TargetReference, $"PivotTable {artifact.Name} target_reference");
        }
        catch (CodecException)
        {
            return false;
        }
        return true;
    }

    private static bool TryReadItemFilter(
        XElement pivotField,
        XElement cacheField,
        string field,
        out SpreadsheetPivotItemFilterArtifact? filter)
    {
        filter = null;
        var itemContainers = pivotField.Elements(Main + "items").ToArray();
        if (itemContainers.Length == 0) return pivotField.Attribute("includeNewItemsInFilter") is null;
        if (itemContainers.Length != 1 || !TryReadSharedItems(cacheField, out var shared)) return false;
        var nativeItems = itemContainers[0].Elements(Main + "item").ToArray();
        if (itemContainers[0].Attribute("count") is { } itemCountAttribute &&
            (!uint.TryParse(itemCountAttribute.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var itemCount) || itemCount != nativeItems.Length)) return false;
        var hidden = new HashSet<int>();
        var seen = new HashSet<int>();
        foreach (var item in nativeItems)
        {
            if (!TryReadOptionalBoolean(item.Attribute("h"), out var isHidden)) return false;
            if (!int.TryParse(item.Attribute("x")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var index))
            {
                if (isHidden == true) return false;
                continue;
            }
            if (index < 0 || index >= shared.Count || !seen.Add(index)) return false;
            if (isHidden == true) hidden.Add(index);
        }
        if (seen.Count != shared.Count || !TryReadOptionalBoolean(pivotField.Attribute("includeNewItemsInFilter"), out var includeNewItems)) return false;
        if (hidden.Count == 0 && includeNewItems != false) return true;

        var visible = Enumerable.Range(0, shared.Count).Where(index => !hidden.Contains(index)).ToArray();
        if (visible.Length == 0) return false;
        SpreadsheetPivotItemFilterMode mode;
        IReadOnlyList<int> selected;
        if (includeNewItems == false)
        {
            mode = SpreadsheetPivotItemFilterMode.Include;
            selected = visible;
        }
        else if (includeNewItems == true)
        {
            mode = SpreadsheetPivotItemFilterMode.Exclude;
            selected = hidden.OrderBy(index => index).ToArray();
        }
        else if (visible.Length < hidden.Count)
        {
            mode = SpreadsheetPivotItemFilterMode.Include;
            selected = visible;
        }
        else
        {
            mode = SpreadsheetPivotItemFilterMode.Exclude;
            selected = hidden.OrderBy(index => index).ToArray();
        }
        if (selected.Count is < 1 or > MaxFilterItems) return false;
        filter = new SpreadsheetPivotItemFilterArtifact { Field = field, Mode = mode };
        filter.Items.Add(selected.Select(index => FilterItem(shared[index])));
        return true;
    }

    private static bool TryReadSharedItems(XElement cacheField, out IReadOnlyList<CacheValue> values)
    {
        values = [];
        var containers = cacheField.Elements(Main + "sharedItems").ToArray();
        if (containers.Length != 1) return false;
        var output = new List<CacheValue>();
        foreach (var item in containers[0].Elements())
        {
            var value = item.Attribute("v")?.Value ?? "";
            if (item.Name == Main + "s") output.Add(new CacheValue(CacheValueKind.String, value, 0, false));
            else if (item.Name == Main + "n" && double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) && double.IsFinite(number))
                output.Add(new CacheValue(CacheValueKind.Number, "", number, false));
            else if (item.Name == Main + "b" && value is "0" or "1") output.Add(new CacheValue(CacheValueKind.Boolean, "", 0, value == "1"));
            else if (item.Name == Main + "e" && !string.IsNullOrEmpty(value)) output.Add(new CacheValue(CacheValueKind.Error, value, 0, false));
            else if (item.Name == Main + "m") output.Add(new CacheValue(CacheValueKind.Blank, "", 0, false));
            else return false;
        }
        if (containers[0].Attribute("count") is { } countAttribute &&
            (!uint.TryParse(countAttribute.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var count) || count != output.Count)) return false;
        values = output;
        return true;
    }

    private static SpreadsheetPivotItemArtifact FilterItem(CacheValue value)
    {
        var output = new SpreadsheetPivotItemArtifact();
        switch (value.Kind)
        {
            case CacheValueKind.String: output.StringValue = value.Text; break;
            case CacheValueKind.Number: output.NumberValue = value.Number; break;
            case CacheValueKind.Boolean: output.BoolValue = value.Boolean; break;
            case CacheValueKind.Blank: output.BlankValue = true; break;
            case CacheValueKind.Error: output.ErrorValue = value.Text; break;
        }
        return output;
    }

    private static string BuildCacheDefinition(
        SpreadsheetPivotTableArtifact pivot,
        string sourceSheetName,
        IReadOnlyList<string> headers,
        IReadOnlyList<CacheValue>[] shared,
        int recordCount,
        string? recordsRelationshipId)
    {
        var policy = pivot.RefreshPolicy ?? new SpreadsheetPivotRefreshPolicyArtifact();
        var root = new XElement(Main + "pivotCacheDefinition",
            new XAttribute(XNamespace.Xmlns + "r", Relationships),
            new XAttribute("saveData", policy.SaveData ? "1" : "0"),
            new XAttribute("refreshOnLoad", policy.RefreshOnLoad ? "1" : "0"),
            new XAttribute("enableRefresh", policy.EnableRefresh ? "1" : "0"),
            new XAttribute("invalid", policy.Invalid ? "1" : "0"),
            new XAttribute("missingItemsLimit", policy.MissingItemsLimit),
            new XAttribute("recordCount", recordCount));
        if (!string.IsNullOrEmpty(recordsRelationshipId)) root.Add(new XAttribute(Relationships + "id", recordsRelationshipId));
        if (!string.IsNullOrEmpty(policy.RefreshedBy)) root.Add(new XAttribute("refreshedBy", policy.RefreshedBy));
        if (!string.IsNullOrEmpty(policy.RefreshedDateIso)) root.Add(new XAttribute("refreshedDateIso", policy.RefreshedDateIso));
        root.Add(new XElement(Main + "cacheSource", new XAttribute("type", "worksheet"),
            new XElement(Main + "worksheetSource", new XAttribute("ref", pivot.SourceReference), new XAttribute("sheet", sourceSheetName))));
        root.Add(new XElement(Main + "cacheFields", new XAttribute("count", headers.Count), headers.Select((header, index) =>
            new XElement(Main + "cacheField", new XAttribute("name", header), new XAttribute("numFmtId", 0), BuildSharedItems(shared[index])))));
        return root.ToString(SaveOptions.DisableFormatting);
    }

    private static XElement BuildSharedItems(IReadOnlyList<CacheValue> values)
    {
        var numbers = values.Where(value => value.Kind == CacheValueKind.Number).Select(value => value.Number).ToArray();
        var root = new XElement(Main + "sharedItems",
            new XAttribute("count", values.Count),
            new XAttribute("containsBlank", values.Any(value => value.Kind == CacheValueKind.Blank) ? "1" : "0"),
            new XAttribute("containsString", values.Any(value => value.Kind is CacheValueKind.String or CacheValueKind.Error) ? "1" : "0"),
            new XAttribute("containsNumber", numbers.Length > 0 ? "1" : "0"),
            new XAttribute("containsInteger", numbers.Length > 0 && numbers.All(number => number == Math.Truncate(number)) ? "1" : "0"));
        if (numbers.Length > 0)
        {
            root.Add(new XAttribute("minValue", numbers.Min().ToString("R", CultureInfo.InvariantCulture)));
            root.Add(new XAttribute("maxValue", numbers.Max().ToString("R", CultureInfo.InvariantCulture)));
        }
        root.Add(values.Select(CacheItem));
        return root;
    }

    private static string BuildCacheRecords(IReadOnlyList<CacheValue[]> rows, IReadOnlyList<CacheValue>[] shared)
    {
        var indexes = shared.Select(values => values.Select((value, index) => (value.Key, index)).ToDictionary(item => item.Key, item => item.index, StringComparer.Ordinal)).ToArray();
        var root = new XElement(Main + "pivotCacheRecords", new XAttribute("count", rows.Count));
        foreach (var row in rows)
            root.Add(new XElement(Main + "r", row.Select((value, column) => new XElement(Main + "x", new XAttribute("v", indexes[column][value.Key])))));
        return root.ToString(SaveOptions.DisableFormatting);
    }

    private static string BuildPivotTableDefinition(
        SpreadsheetPivotTableArtifact pivot,
        IReadOnlyList<string> headers,
        IReadOnlyList<CacheValue>[] shared,
        IReadOnlyList<int> rowIndexes,
        int columnIndex,
        IReadOnlyList<int> valueIndexes,
        IReadOnlyDictionary<int, ResolvedItemFilter> filters,
        IReadOnlyList<IReadOnlyList<int>> rowItemIndexes,
        IReadOnlyList<int> columnItemIndexes,
        uint cacheId,
        RangeBounds targetRange)
    {
        var valueIndexSet = valueIndexes.ToHashSet();
        var rowIndexSet = rowIndexes.ToHashSet();
        var axisIndexSet = rowIndexes.Append(columnIndex).Where(index => index >= 0).ToHashSet();
        var multipleValues = valueIndexes.Count > 1;
        var multipleRowFields = rowIndexes.Count > 1;
        var root = new XElement(Main + "pivotTableDefinition",
            new XAttribute("name", pivot.Name),
            new XAttribute("cacheId", cacheId),
            new XAttribute("dataCaption", "Values"),
            new XAttribute("dataOnRows", "0"),
            new XAttribute("showDrill", "1"),
            new XAttribute("rowGrandTotals", pivot.RowGrandTotals ? "1" : "0"),
            new XAttribute("colGrandTotals", pivot.ColumnGrandTotals ? "1" : "0"),
            new XAttribute("compact", multipleRowFields ? "0" : "1"),
            new XAttribute("outline", "0"));
        root.Add(new XElement(Main + "location", new XAttribute("ref", RangeAddress(targetRange)), new XAttribute("firstHeaderRow", 0), new XAttribute("firstDataRow", 1), new XAttribute("firstDataCol", rowIndexes.Count)));
        root.Add(new XElement(Main + "pivotFields", new XAttribute("count", headers.Count), headers.Select((header, index) =>
        {
            filters.TryGetValue(index, out var filter);
            var field = new XElement(Main + "pivotField", new XAttribute("name", header), new XAttribute("dataField", valueIndexSet.Contains(index) ? "1" : "0"), new XAttribute("subtotalTop", "1"), new XAttribute("showAll", filters.Count > 0 && axisIndexSet.Contains(index) ? "0" : "1"));
            if (rowIndexSet.Contains(index)) field.Add(new XAttribute("axis", "axisRow"));
            else if (index == columnIndex) field.Add(new XAttribute("axis", "axisCol"));
            if (multipleRowFields && rowIndexSet.Contains(index))
            {
                field.Add(new XAttribute("compact", "0"));
                field.Add(new XAttribute("outline", "0"));
                field.Add(new XAttribute("defaultSubtotal", "0"));
            }
            if (filter is not null)
            {
                field.Add(new XAttribute("multipleItemSelectionAllowed", "1"));
                field.Add(new XAttribute("includeNewItemsInFilter", filter.Mode == SpreadsheetPivotItemFilterMode.Exclude ? "1" : "0"));
            }
            if (axisIndexSet.Contains(index))
                field.Add(new XElement(Main + "items", new XAttribute("count", shared[index].Count), shared[index].Select((item, itemIndex) =>
                {
                    var element = new XElement(Main + "item", new XAttribute("x", itemIndex));
                    if (filter is not null && !filter.Visible(item)) element.Add(new XAttribute("h", "1"));
                    return element;
                })));
            return field;
        })));
        root.Add(new XElement(Main + "rowFields", new XAttribute("count", rowIndexes.Count), rowIndexes.Select(index => new XElement(Main + "field", new XAttribute("x", index)))));
        root.Add(new XElement(Main + "rowItems", new XAttribute("count", rowItemIndexes.Count + (pivot.ColumnGrandTotals ? 1 : 0)),
            rowItemIndexes.Select(tuple => new XElement(Main + "i", tuple.Select(index => new XElement(Main + "x", new XAttribute("v", index))))),
            pivot.ColumnGrandTotals ? (multipleRowFields
                ? new XElement(Main + "i", new XAttribute("t", "grand"), rowIndexes.Select(_ => new XElement(Main + "x")))
                : new XElement(Main + "i")) : null));
        if (columnIndex >= 0 || multipleValues)
        {
            var columnFields = new List<XElement>();
            if (columnIndex >= 0) columnFields.Add(new XElement(Main + "field", new XAttribute("x", columnIndex)));
            if (multipleValues) columnFields.Add(new XElement(Main + "field", new XAttribute("x", DataLayoutFieldIndex)));
            root.Add(new XElement(Main + "colFields", new XAttribute("count", columnFields.Count), columnFields));

            if (!multipleValues)
            {
                root.Add(new XElement(Main + "colItems", new XAttribute("count", columnItemIndexes.Count + (pivot.RowGrandTotals ? 1 : 0)),
                    columnItemIndexes.Select(index => new XElement(Main + "i", new XElement(Main + "x", new XAttribute("v", index)))),
                    pivot.RowGrandTotals ? new XElement(Main + "i") : null));
            }
            else
            {
                var columnItems = new List<XElement>();
                var categoryIndexes = columnIndex >= 0 ? columnItemIndexes : [0];
                foreach (var categoryIndex in categoryIndexes)
                    for (var valueOrdinal = 0; valueOrdinal < valueIndexes.Count; valueOrdinal++)
                    {
                        var item = new XElement(Main + "i", new XAttribute("i", valueOrdinal));
                        if (columnIndex >= 0) item.Add(new XElement(Main + "x", new XAttribute("v", categoryIndex)));
                        item.Add(new XElement(Main + "x", new XAttribute("v", valueOrdinal)));
                        columnItems.Add(item);
                    }
                if (columnIndex >= 0 && pivot.RowGrandTotals)
                    for (var valueOrdinal = 0; valueOrdinal < valueIndexes.Count; valueOrdinal++)
                    {
                        var item = new XElement(Main + "i", new XAttribute("t", "grand"), new XAttribute("i", valueOrdinal), new XElement(Main + "x"));
                        item.Add(new XElement(Main + "x", new XAttribute("v", valueOrdinal)));
                        columnItems.Add(item);
                    }
                root.Add(new XElement(Main + "colItems", new XAttribute("count", columnItems.Count), columnItems));
            }
        }
        root.Add(new XElement(Main + "dataFields", new XAttribute("count", pivot.ValueFields.Count), pivot.ValueFields.Select((value, index) =>
            new XElement(Main + "dataField",
                new XAttribute("name", string.IsNullOrEmpty(value.Name) ? $"{AggregationLabel(value.Aggregation)} of {value.Field}" : value.Name),
                new XAttribute("fld", valueIndexes[index]), new XAttribute("subtotal", AggregationName(value.Aggregation))))));
        root.Add(new XElement(Main + "pivotTableStyleInfo", new XAttribute("showRowHeaders", "1"), new XAttribute("showColHeaders", "1"), new XAttribute("showRowStripes", "0"), new XAttribute("showColStripes", "0"), new XAttribute("showLastColumn", "0")));
        return root.ToString(SaveOptions.DisableFormatting);
    }

    private static void ValidateSemantic(SpreadsheetPivotTableArtifact pivot)
    {
        if (string.IsNullOrWhiteSpace(pivot.Id) || string.IsNullOrWhiteSpace(pivot.Name) || pivot.Name.Length > 255) throw Invalid("PivotTable id and a 1-255 character name are required.");
        if (pivot.RowFields.Count is < 1 or > MaxRowFields || pivot.ColumnFields.Count > 1 || pivot.ValueFields.Count is < 1 or > MaxValueFields)
            throw new CodecException("unsupported_spreadsheet_pivot_profile", $"PivotTable {pivot.Name} requires 1 through {MaxRowFields} row fields, at most one column field, and 1 through {MaxValueFields} value fields.");
        var axisFields = pivot.RowFields.Concat(pivot.ColumnFields).ToArray();
        if (axisFields.Any(string.IsNullOrWhiteSpace) || axisFields.Distinct(StringComparer.Ordinal).Count() != axisFields.Length)
            throw Invalid($"PivotTable {pivot.Name} axis fields must be non-empty, unique, and cannot appear on both axes.");
        if (pivot.ValueFields.Any(value => !Enum.IsDefined(value.Aggregation) || value.Aggregation == SpreadsheetPivotAggregation.Unspecified))
            throw Invalid($"PivotTable {pivot.Name} has an unsupported aggregation.");
        if (pivot.ItemFilters.Count > axisFields.Length || pivot.ItemFilters.Select(filter => filter.Field).Distinct(StringComparer.Ordinal).Count() != pivot.ItemFilters.Count)
            throw new CodecException("unsupported_spreadsheet_pivot_filter", $"PivotTable {pivot.Name} requires at most one item filter per native axis field.");
        var axisFieldSet = axisFields.ToHashSet(StringComparer.Ordinal);
        foreach (var filter in pivot.ItemFilters)
        {
            if (!axisFieldSet.Contains(filter.Field) || filter.Mode is not (SpreadsheetPivotItemFilterMode.Include or SpreadsheetPivotItemFilterMode.Exclude) ||
                filter.Items.Count is < 1 or > MaxFilterItems)
                throw new CodecException("unsupported_spreadsheet_pivot_filter", $"PivotTable {pivot.Name} item filter {filter.Field} is outside the bounded native profile.");
            if (filter.Items.Select(FilterItemKey).Distinct(StringComparer.Ordinal).Count() != filter.Items.Count)
                throw Invalid($"PivotTable {pivot.Name} item filter {filter.Field} contains duplicate items.");
        }
        if (string.IsNullOrWhiteSpace(pivot.SourceWorksheetId) || string.IsNullOrWhiteSpace(pivot.SourceReference) || string.IsNullOrWhiteSpace(pivot.TargetReference)) throw Invalid($"PivotTable {pivot.Name} requires source and target references.");
    }

    private static IReadOnlyDictionary<int, ResolvedItemFilter> ResolveItemFilters(
        SpreadsheetPivotTableArtifact pivot,
        IReadOnlyList<string> headers,
        IReadOnlyList<CacheValue>[] shared)
    {
        var output = new Dictionary<int, ResolvedItemFilter>();
        foreach (var filter in pivot.ItemFilters)
        {
            var fieldIndex = -1;
            for (var index = 0; index < headers.Count; index++)
                if (string.Equals(headers[index], filter.Field, StringComparison.Ordinal))
                {
                    fieldIndex = index;
                    break;
                }
            if (fieldIndex < 0) throw Invalid($"PivotTable {pivot.Name} item filter field {filter.Field} is absent from its source headers.");
            var keys = filter.Items.Select(FilterItemKey).ToHashSet(StringComparer.Ordinal);
            var available = shared[fieldIndex].Select(item => item.Key).ToHashSet(StringComparer.Ordinal);
            if (!keys.IsSubsetOf(available)) throw Invalid($"PivotTable {pivot.Name} item filter {filter.Field} references an item outside its source cache.");
            var resolved = new ResolvedItemFilter(filter.Mode, keys);
            if (!shared[fieldIndex].Any(resolved.Visible))
                throw new CodecException("unsupported_spreadsheet_pivot_filter", $"PivotTable {pivot.Name} item filter {filter.Field} hides every field item.");
            output.Add(fieldIndex, resolved);
        }
        return output;
    }

    private static IReadOnlyList<int> ActiveItemIndexes(
        IEnumerable<CacheValue[]> rows,
        int fieldIndex,
        IReadOnlyList<CacheValue> shared)
    {
        var indexes = shared.Select((item, index) => (item.Key, index)).ToDictionary(item => item.Key, item => item.index, StringComparer.Ordinal);
        var output = new List<int>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var row in rows)
            if (seen.Add(row[fieldIndex].Key)) output.Add(indexes[row[fieldIndex].Key]);
        return output;
    }

    private static IReadOnlyList<IReadOnlyList<int>> ActiveItemTuples(
        IEnumerable<CacheValue[]> rows,
        IReadOnlyList<int> fieldIndexes,
        IReadOnlyList<CacheValue>[] shared)
    {
        var indexes = fieldIndexes.Select(fieldIndex => shared[fieldIndex]
            .Select((item, index) => (item.Key, index))
            .ToDictionary(item => item.Key, item => item.index, StringComparer.Ordinal)).ToArray();
        var output = new List<IReadOnlyList<int>>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var row in rows)
        {
            var keys = fieldIndexes.Select(fieldIndex => row[fieldIndex].Key).ToArray();
            var tupleKey = string.Concat(keys.Select(key => $"{key.Length}:{key}"));
            if (seen.Add(tupleKey)) output.Add(keys.Select((key, ordinal) => indexes[ordinal][key]).ToArray());
        }
        return output;
    }

    private static string FilterItemKey(SpreadsheetPivotItemArtifact item) => item.ValueCase switch
    {
        SpreadsheetPivotItemArtifact.ValueOneofCase.StringValue => $"s:{item.StringValue}",
        SpreadsheetPivotItemArtifact.ValueOneofCase.NumberValue when double.IsFinite(item.NumberValue) => $"n:{item.NumberValue.ToString("R", CultureInfo.InvariantCulture)}",
        SpreadsheetPivotItemArtifact.ValueOneofCase.BoolValue => item.BoolValue ? "b:1" : "b:0",
        SpreadsheetPivotItemArtifact.ValueOneofCase.BlankValue when item.BlankValue => "m:",
        SpreadsheetPivotItemArtifact.ValueOneofCase.ErrorValue when !string.IsNullOrEmpty(item.ErrorValue) => $"e:{item.ErrorValue}",
        _ => throw new CodecException("unsupported_spreadsheet_pivot_filter", "PivotTable item filter contains an unsupported cached value."),
    };

    private static IReadOnlyList<CacheValue> DistinctValues(IEnumerable<CacheValue> values)
    {
        var output = new List<CacheValue>();
        var keys = new HashSet<string>(StringComparer.Ordinal);
        foreach (var value in values) if (keys.Add(value.Key)) output.Add(value);
        return output;
    }

    private static XElement CacheItem(CacheValue value) => value.Kind switch
    {
        CacheValueKind.Blank => new XElement(Main + "m"),
        CacheValueKind.Number => new XElement(Main + "n", new XAttribute("v", value.Number.ToString("R", CultureInfo.InvariantCulture))),
        CacheValueKind.Boolean => new XElement(Main + "b", new XAttribute("v", value.Boolean ? "1" : "0")),
        CacheValueKind.Error => new XElement(Main + "e", new XAttribute("v", value.Text)),
        _ => new XElement(Main + "s", new XAttribute("v", value.Text)),
    };

    private static string CellText(CellArtifact? cell) => cell?.ValueCase switch
    {
        CellArtifact.ValueOneofCase.StringValue => cell.StringValue,
        CellArtifact.ValueOneofCase.NumberValue => cell.NumberValue.ToString("R", CultureInfo.InvariantCulture),
        CellArtifact.ValueOneofCase.BoolValue => cell.BoolValue ? "TRUE" : "FALSE",
        CellArtifact.ValueOneofCase.ErrorValue => cell.ErrorValue,
        _ => string.Empty,
    };

    private static bool TryAggregation(string? name, out SpreadsheetPivotAggregation aggregation)
    {
        aggregation = name switch
        {
            null or "sum" => SpreadsheetPivotAggregation.Sum,
            "count" => SpreadsheetPivotAggregation.Count,
            "average" or "avg" => SpreadsheetPivotAggregation.Average,
            "min" => SpreadsheetPivotAggregation.Min,
            "max" => SpreadsheetPivotAggregation.Max,
            _ => SpreadsheetPivotAggregation.Unspecified,
        };
        return aggregation != SpreadsheetPivotAggregation.Unspecified;
    }

    private static string AggregationName(SpreadsheetPivotAggregation aggregation) => aggregation switch
    {
        SpreadsheetPivotAggregation.Sum => "sum",
        SpreadsheetPivotAggregation.Count => "count",
        SpreadsheetPivotAggregation.Average => "average",
        SpreadsheetPivotAggregation.Min => "min",
        SpreadsheetPivotAggregation.Max => "max",
        _ => throw Invalid("PivotTable aggregation is unsupported."),
    };

    private static string AggregationLabel(SpreadsheetPivotAggregation aggregation) => aggregation switch
    {
        SpreadsheetPivotAggregation.Sum => "Sum",
        SpreadsheetPivotAggregation.Count => "Count",
        SpreadsheetPivotAggregation.Average => "Average",
        SpreadsheetPivotAggregation.Min => "Min",
        SpreadsheetPivotAggregation.Max => "Max",
        _ => "Value",
    };

    private static RangeBounds ParseRange(string reference, string label)
    {
        var match = A1Range.Match(reference ?? string.Empty);
        if (!match.Success) throw Invalid($"{label} must be a bounded A1 cell or rectangle.");
        var left = ParseColumn(match.Groups["firstColumn"].Value);
        var top = uint.Parse(match.Groups["firstRow"].Value, CultureInfo.InvariantCulture) - 1;
        var right = match.Groups["lastColumn"].Success ? ParseColumn(match.Groups["lastColumn"].Value) : left;
        var bottom = match.Groups["lastRow"].Success ? uint.Parse(match.Groups["lastRow"].Value, CultureInfo.InvariantCulture) - 1 : top;
        if (left > right || top > bottom || right > 16_383 || bottom > 1_048_575) throw Invalid($"{label} is outside the XLSX worksheet grid.");
        return new RangeBounds(top, left, bottom, right);
    }

    private static uint ParseColumn(string label)
    {
        uint value = 0;
        foreach (var character in label.ToUpperInvariant()) value = checked(value * 26 + (uint)(character - 'A' + 1));
        return value - 1;
    }

    private static string ColumnLabel(uint column)
    {
        Span<char> buffer = stackalloc char[3];
        var position = buffer.Length;
        var value = column + 1;
        while (value > 0)
        {
            value--;
            buffer[--position] = (char)('A' + value % 26);
            value /= 26;
        }
        return new string(buffer[position..]);
    }

    private static string RangeAddress(RangeBounds bounds)
    {
        var first = $"{ColumnLabel(bounds.Left)}{bounds.Top + 1}";
        var last = $"{ColumnLabel(bounds.Right)}{bounds.Bottom + 1}";
        return first == last ? first : $"{first}:{last}";
    }

    private static bool ReadBoolean(XAttribute? attribute, bool fallback) => attribute?.Value switch
    {
        "1" or "true" or "TRUE" => true,
        "0" or "false" or "FALSE" => false,
        null => fallback,
        _ => fallback,
    };

    private static bool TryReadOptionalBoolean(XAttribute? attribute, out bool? value)
    {
        value = attribute?.Value switch
        {
            "1" or "true" or "TRUE" => true,
            "0" or "false" or "FALSE" => false,
            null => null,
            _ => null,
        };
        return attribute is null || attribute.Value is "1" or "true" or "TRUE" or "0" or "false" or "FALSE";
    }

    private static uint ReadUInt(XAttribute? attribute) => uint.TryParse(attribute?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var value) ? value : 0;

    private static string SemanticHash(SpreadsheetPivotTableArtifact pivot)
    {
        var policy = pivot.RefreshPolicy ?? new SpreadsheetPivotRefreshPolicyArtifact();
        return Hash(string.Join("\u001f", pivot.Id, pivot.Name, pivot.SourceWorksheetId, pivot.SourceReference, pivot.TargetReference,
            string.Join("\u001e", pivot.RowFields), string.Join("\u001e", pivot.ColumnFields),
            string.Join("\u001d", pivot.ValueFields.Select(value => string.Join("\u001c", value.Field, value.Name, (int)value.Aggregation))),
            string.Join("\u001d", pivot.ItemFilters.Select(filter => string.Join("\u001c", filter.Field, (int)filter.Mode, string.Join("\u001b", filter.Items.Select(FilterItemKey))))),
            pivot.RowGrandTotals, pivot.ColumnGrandTotals, policy.RefreshOnLoad, policy.SaveData, policy.EnableRefresh, policy.Invalid,
            policy.MissingItemsLimit, policy.RefreshedBy, policy.RefreshedDateIso));
    }

    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static string Hash(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();
    private static CodecException Invalid(string message) => new("invalid_spreadsheet_pivot", message);

    private sealed record PivotRecord(
        SpreadsheetPivotTableArtifact Artifact,
        PivotTablePart PivotPart,
        PivotTableCacheDefinitionPart CachePart,
        PivotTableCacheRecordsPart? RecordsPart,
        uint Ordinal,
        uint CacheId);

    private sealed record ResolvedItemFilter(SpreadsheetPivotItemFilterMode Mode, HashSet<string> Keys)
    {
        internal bool Visible(CacheValue value) => Mode == SpreadsheetPivotItemFilterMode.Include
            ? Keys.Contains(value.Key)
            : !Keys.Contains(value.Key);
    }

    private readonly record struct RangeBounds(uint Top, uint Left, uint Bottom, uint Right)
    {
        internal int RowCount => checked((int)(Bottom - Top + 1));
        internal int ColumnCount => checked((int)(Right - Left + 1));
    }

    private enum CacheValueKind { Blank, String, Number, Boolean, Error }

    private readonly record struct CacheValue(CacheValueKind Kind, string Text, double Number, bool Boolean)
    {
        internal string Key => Kind switch
        {
            CacheValueKind.Number => $"n:{Number.ToString("R", CultureInfo.InvariantCulture)}",
            CacheValueKind.Boolean => Boolean ? "b:1" : "b:0",
            CacheValueKind.Blank => "m:",
            CacheValueKind.Error => $"e:{Text}",
            _ => $"s:{Text}",
        };

        internal static CacheValue FromCell(CellArtifact? cell) => cell?.ValueCase switch
        {
            CellArtifact.ValueOneofCase.StringValue => new(CacheValueKind.String, cell.StringValue, 0, false),
            CellArtifact.ValueOneofCase.NumberValue when double.IsFinite(cell.NumberValue) => new(CacheValueKind.Number, "", cell.NumberValue, false),
            CellArtifact.ValueOneofCase.BoolValue => new(CacheValueKind.Boolean, "", 0, cell.BoolValue),
            CellArtifact.ValueOneofCase.ErrorValue => new(CacheValueKind.Error, cell.ErrorValue, 0, false),
            _ => new(CacheValueKind.Blank, "", 0, false),
        };
    }
}
