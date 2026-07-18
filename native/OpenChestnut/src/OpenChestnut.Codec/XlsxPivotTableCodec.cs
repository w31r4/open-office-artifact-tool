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
        var rowIndex = Array.IndexOf(headers, pivot.RowFields[0]);
        var columnIndex = pivot.ColumnFields.Count == 0 ? -1 : Array.IndexOf(headers, pivot.ColumnFields[0]);
        var valueIndex = Array.IndexOf(headers, pivot.ValueFields[0].Field);
        if (rowIndex < 0 || (pivot.ColumnFields.Count > 0 && columnIndex < 0) || valueIndex < 0) throw Invalid($"PivotTable {pivot.Name} contains a field outside its source headers.");

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

        var expectedRows = 1 + shared[rowIndex].Count + (pivot.ColumnGrandTotals ? 1 : 0);
        var expectedColumns = 1 + (columnIndex >= 0 ? shared[columnIndex].Count + (pivot.RowGrandTotals ? 1 : 0) : 1);
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
        pivotPart.PivotTableDefinition = new PivotTableDefinition(BuildPivotTableDefinition(pivot, headers, shared, rowIndex, columnIndex, valueIndex, cacheId, targetRange));

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
        var rowField = root.Elements(Main + "rowFields").SingleOrDefault()?.Elements(Main + "field").SingleOrDefault();
        var columnFieldElements = root.Elements(Main + "colFields").SingleOrDefault()?.Elements(Main + "field").ToArray() ?? [];
        var dataField = root.Elements(Main + "dataFields").SingleOrDefault()?.Elements(Main + "dataField").SingleOrDefault();
        var cacheFields = cacheRoot.Elements(Main + "cacheFields").SingleOrDefault()?.Elements(Main + "cacheField").ToArray();
        var worksheetSource = cacheRoot.Elements(Main + "cacheSource").SingleOrDefault()?.Elements(Main + "worksheetSource").SingleOrDefault();
        if (location is null || pivotFields is null || rowField is null || columnFieldElements.Length > 1 || dataField is null || cacheFields is null || worksheetSource is null || pivotFields.Length != cacheFields.Length) return false;
        if (!int.TryParse(rowField.Attribute("x")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var rowIndex) || rowIndex < 0 || rowIndex >= cacheFields.Length) return false;
        var columnIndex = -1;
        if (columnFieldElements.Length == 1 && (!int.TryParse(columnFieldElements[0].Attribute("x")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out columnIndex) || columnIndex < 0 || columnIndex >= cacheFields.Length)) return false;
        if (!int.TryParse(dataField.Attribute("fld")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var valueIndex) || valueIndex < 0 || valueIndex >= cacheFields.Length) return false;
        if (!TryAggregation(dataField.Attribute("subtotal")?.Value, out var aggregation)) return false;
        var headers = cacheFields.Select(field => field.Attribute("name")?.Value ?? "").ToArray();
        if (headers.Any(string.IsNullOrWhiteSpace) || headers.Distinct(StringComparer.Ordinal).Count() != headers.Length) return false;
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
        artifact.RowFields.Add(headers[rowIndex]);
        if (columnIndex >= 0) artifact.ColumnFields.Add(headers[columnIndex]);
        artifact.ValueFields.Add(new SpreadsheetPivotValueFieldArtifact
        {
            Field = headers[valueIndex],
            Name = dataField.Attribute("name")?.Value ?? "",
            Aggregation = aggregation,
        });
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
        int rowIndex,
        int columnIndex,
        int valueIndex,
        uint cacheId,
        RangeBounds targetRange)
    {
        var root = new XElement(Main + "pivotTableDefinition",
            new XAttribute("name", pivot.Name),
            new XAttribute("cacheId", cacheId),
            new XAttribute("dataCaption", "Values"),
            new XAttribute("dataOnRows", "0"),
            new XAttribute("showDrill", "1"),
            new XAttribute("rowGrandTotals", pivot.RowGrandTotals ? "1" : "0"),
            new XAttribute("colGrandTotals", pivot.ColumnGrandTotals ? "1" : "0"),
            new XAttribute("compact", "1"),
            new XAttribute("outline", "0"));
        root.Add(new XElement(Main + "location", new XAttribute("ref", RangeAddress(targetRange)), new XAttribute("firstHeaderRow", 0), new XAttribute("firstDataRow", 1), new XAttribute("firstDataCol", 1)));
        root.Add(new XElement(Main + "pivotFields", new XAttribute("count", headers.Count), headers.Select((header, index) =>
        {
            var field = new XElement(Main + "pivotField", new XAttribute("name", header), new XAttribute("dataField", index == valueIndex ? "1" : "0"), new XAttribute("subtotalTop", "1"), new XAttribute("showAll", "1"));
            if (index == rowIndex) field.Add(new XAttribute("axis", "axisRow"));
            else if (index == columnIndex) field.Add(new XAttribute("axis", "axisCol"));
            if (index == rowIndex || index == columnIndex)
                field.Add(new XElement(Main + "items", new XAttribute("count", shared[index].Count), shared[index].Select((_, itemIndex) => new XElement(Main + "item", new XAttribute("x", itemIndex)))));
            return field;
        })));
        root.Add(new XElement(Main + "rowFields", new XAttribute("count", 1), new XElement(Main + "field", new XAttribute("x", rowIndex))));
        root.Add(new XElement(Main + "rowItems", new XAttribute("count", shared[rowIndex].Count + (pivot.ColumnGrandTotals ? 1 : 0)),
            shared[rowIndex].Select((_, index) => new XElement(Main + "i", new XElement(Main + "x", new XAttribute("v", index)))),
            pivot.ColumnGrandTotals ? new XElement(Main + "i") : null));
        if (columnIndex >= 0)
        {
            root.Add(new XElement(Main + "colFields", new XAttribute("count", 1), new XElement(Main + "field", new XAttribute("x", columnIndex))));
            root.Add(new XElement(Main + "colItems", new XAttribute("count", shared[columnIndex].Count + (pivot.RowGrandTotals ? 1 : 0)),
                shared[columnIndex].Select((_, index) => new XElement(Main + "i", new XElement(Main + "x", new XAttribute("v", index)))),
                pivot.RowGrandTotals ? new XElement(Main + "i") : null));
        }
        var value = pivot.ValueFields[0];
        root.Add(new XElement(Main + "dataFields", new XAttribute("count", 1), new XElement(Main + "dataField",
            new XAttribute("name", string.IsNullOrEmpty(value.Name) ? $"{AggregationLabel(value.Aggregation)} of {value.Field}" : value.Name),
            new XAttribute("fld", valueIndex), new XAttribute("subtotal", AggregationName(value.Aggregation)))));
        root.Add(new XElement(Main + "pivotTableStyleInfo", new XAttribute("showRowHeaders", "1"), new XAttribute("showColHeaders", "1"), new XAttribute("showRowStripes", "0"), new XAttribute("showColStripes", "0"), new XAttribute("showLastColumn", "0")));
        return root.ToString(SaveOptions.DisableFormatting);
    }

    private static void ValidateSemantic(SpreadsheetPivotTableArtifact pivot)
    {
        if (string.IsNullOrWhiteSpace(pivot.Id) || string.IsNullOrWhiteSpace(pivot.Name) || pivot.Name.Length > 255) throw Invalid("PivotTable id and a 1-255 character name are required.");
        if (pivot.RowFields.Count != 1 || pivot.ColumnFields.Count > 1 || pivot.ValueFields.Count != 1) throw new CodecException("unsupported_spreadsheet_pivot_profile", $"PivotTable {pivot.Name} requires one row field, at most one column field, and one value field.");
        if (pivot.RowFields[0] == pivot.ColumnFields.FirstOrDefault()) throw Invalid($"PivotTable {pivot.Name} cannot place one field on both axes.");
        if (!Enum.IsDefined(pivot.ValueFields[0].Aggregation) || pivot.ValueFields[0].Aggregation == SpreadsheetPivotAggregation.Unspecified) throw Invalid($"PivotTable {pivot.Name} has an unsupported aggregation.");
        if (string.IsNullOrWhiteSpace(pivot.SourceWorksheetId) || string.IsNullOrWhiteSpace(pivot.SourceReference) || string.IsNullOrWhiteSpace(pivot.TargetReference)) throw Invalid($"PivotTable {pivot.Name} requires source and target references.");
    }

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

    private static uint ReadUInt(XAttribute? attribute) => uint.TryParse(attribute?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var value) ? value : 0;

    private static string SemanticHash(SpreadsheetPivotTableArtifact pivot)
    {
        var value = pivot.ValueFields.Single();
        var policy = pivot.RefreshPolicy ?? new SpreadsheetPivotRefreshPolicyArtifact();
        return Hash(string.Join("\u001f", pivot.Id, pivot.Name, pivot.SourceWorksheetId, pivot.SourceReference, pivot.TargetReference,
            string.Join("\u001e", pivot.RowFields), string.Join("\u001e", pivot.ColumnFields), value.Field, value.Name, (int)value.Aggregation,
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
