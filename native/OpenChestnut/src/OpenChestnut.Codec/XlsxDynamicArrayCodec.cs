using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenOffice.Artifact.Wire.V1;
using DynamicArrayProperties = DocumentFormat.OpenXml.Office2019.Excel.DynamicArray.DynamicArrayProperties;
using SpreadsheetExtension = DocumentFormat.OpenXml.Spreadsheet.Extension;

namespace OpenChestnut.Codec;

// Dynamic arrays use the same worksheet <f t="array"> shape as legacy array
// formulas. Their distinguishing state belongs to the workbook-level
// CellMetadataPart and the anchor cell's one-based cm index. Keeping that
// package graph here prevents formula text handling from guessing at metadata
// indexes or manufacturing a private marker.
internal sealed class XlsxDynamicArrayCodec
{
    private const string MetadataName = "XLDAPR";
    private const string ExtensionUri = "{BDBB8CDC-FA1E-496E-A857-3C3F30C029C3}";
    private const string DynamicArrayNamespace = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray";

    private readonly WorkbookPart _workbookPart;
    private readonly HashSet<uint> _dynamicIndexes = [];
    private readonly HashSet<uint> _exclusiveDynamicIndexes = [];
    private uint? _authoredIndex;

    internal XlsxDynamicArrayCodec(WorkbookPart workbookPart)
    {
        _workbookPart = workbookPart;
        ReadIndexes();
    }

    internal bool IsDynamic(Cell cell) =>
        cell.CellMetaIndex?.Value is { } index && _dynamicIndexes.Contains(index);

    internal bool HasOpaqueMetadata(Cell cell) =>
        cell.CellMetaIndex?.Value is { } index && !_dynamicIndexes.Contains(index);

    internal void ConfigureNewCell(Cell cell, CellArtifact source, bool sourceBound)
    {
        if (source.FormulaMetadata?.Kind != CellFormulaKind.DynamicArray) return;
        cell.CellMetaIndex = EnsureDynamicIndex(sourceBound);
    }

    internal void ApplyFormulaMetadata(Cell cell, CellArtifact desired, bool sourceBound)
    {
        var desiredDynamic = desired.FormulaMetadata?.Kind == CellFormulaKind.DynamicArray;
        var currentIndex = cell.CellMetaIndex?.Value;
        if (desiredDynamic)
        {
            if (currentIndex is not null)
            {
                if (_dynamicIndexes.Contains(currentIndex.Value)) return;
                throw Unsupported(desired, "anchor has non-dynamic cell metadata and cannot be converted losslessly");
            }
            cell.CellMetaIndex = EnsureDynamicIndex(sourceBound);
            return;
        }

        if (currentIndex is null) return;
        if (!_dynamicIndexes.Contains(currentIndex.Value))
            throw Unsupported(desired, "formula is linked to opaque cell metadata and cannot be edited losslessly");
        if (!_exclusiveDynamicIndexes.Contains(currentIndex.Value))
            throw Unsupported(desired, "dynamic-array marker shares its cell metadata block with unmodeled records and cannot be removed losslessly");
        cell.CellMetaIndex = null;
    }

    private uint EnsureDynamicIndex(bool sourceBound)
    {
        if (_authoredIndex is not null) return _authoredIndex.Value;
        if (_dynamicIndexes.Count > 0) return _dynamicIndexes.Min();
        if (sourceBound)
            throw new CodecException(
                "unsupported_dynamic_array_edit",
                "Source-bound workbook has no recognized XLDAPR cell-metadata record; adding a dynamic-array formula would require changing an opaque metadata graph.",
                _workbookPart.Uri.ToString());
        if (_workbookPart.CellMetadataPart is not null)
            throw new CodecException(
                "unsupported_dynamic_array_edit",
                "Workbook already has a non-dynamic CellMetadataPart; OpenChestnut will not replace it while authoring a dynamic array.",
                _workbookPart.CellMetadataPart.Uri.ToString());

        var part = _workbookPart.AddNewPart<CellMetadataPart>();
        var metadataType = new MetadataType
        {
            Name = MetadataName,
            MinSupportedVersion = 120_000U,
            Copy = true,
            PasteAll = true,
            PasteValues = true,
            Merge = true,
            SplitFirst = true,
            RowColumnShift = true,
            ClearFormats = true,
            ClearComments = true,
            Assign = true,
            Coerce = true,
            CellMeta = true,
        };
        var extension = new SpreadsheetExtension(
            new DynamicArrayProperties { FDynamic = true, FCollapsed = false })
        {
            Uri = ExtensionUri,
        };
        part.Metadata = new Metadata(
            new MetadataTypes(metadataType) { Count = 1U },
            new FutureMetadata(
                new FutureMetadataBlock(
                    new ExtensionList(extension)))
            {
                Name = MetadataName,
                Count = 1U,
            },
            new CellMetadata(
                new MetadataBlock(
                    new MetadataRecord { TypeIndex = 1U, Val = 0U }))
            {
                Count = 1U,
            });
        part.Metadata.Save();
        _authoredIndex = 1U;
        _dynamicIndexes.Add(1U);
        _exclusiveDynamicIndexes.Add(1U);
        return 1U;
    }

    private void ReadIndexes()
    {
        var metadata = _workbookPart.CellMetadataPart?.Metadata;
        if (metadata is null) return;
        var types = metadata.GetFirstChild<MetadataTypes>()?.Elements<MetadataType>().ToArray() ?? [];
        var dynamicTypes = types
            .Select((item, index) => (Item: item, Index: checked((uint)index + 1)))
            .Where(item => string.Equals(item.Item.Name?.Value, MetadataName, StringComparison.Ordinal) && item.Item.CellMeta?.Value == true)
            .ToArray();
        if (dynamicTypes.Length != 1) return;
        var dynamicTypeIndex = dynamicTypes[0].Index;

        var future = metadata.Elements<FutureMetadata>()
            .Where(item => string.Equals(item.Name?.Value, MetadataName, StringComparison.Ordinal))
            .ToArray();
        if (future.Length != 1) return;
        var futureBlocks = future[0].Elements<FutureMetadataBlock>().ToArray();
        if (future[0].Count?.Value != (uint)futureBlocks.Length) return;
        var dynamicValues = futureBlocks
            .Select((block, index) => (Block: block, Index: checked((uint)index)))
            .Where(item => IsDynamicBlock(item.Block))
            .Select(item => item.Index)
            .ToHashSet();
        if (dynamicValues.Count == 0) return;

        var cellMetadata = metadata.Elements<CellMetadata>().ToArray();
        if (cellMetadata.Length != 1) return;
        var blocks = cellMetadata[0].Elements<MetadataBlock>().ToArray();
        if (cellMetadata[0].Count?.Value != (uint)blocks.Length) return;
        for (var index = 0; index < blocks.Length; index++)
        {
            var records = blocks[index].Elements<MetadataRecord>().ToArray();
            var dynamicRecords = records.Where(record =>
                record.TypeIndex?.Value == dynamicTypeIndex &&
                record.Val?.Value is { } value && dynamicValues.Contains(value)).ToArray();
            if (dynamicRecords.Length == 0) continue;
            var cellIndex = checked((uint)index + 1);
            _dynamicIndexes.Add(cellIndex);
            if (records.Length == 1 && dynamicRecords.Length == 1)
                _exclusiveDynamicIndexes.Add(cellIndex);
        }
    }

    private static bool IsDynamicBlock(FutureMetadataBlock block)
    {
        var matches = block.Descendants<SpreadsheetExtension>()
            .Where(extension => string.Equals(extension.Uri?.Value, ExtensionUri, StringComparison.OrdinalIgnoreCase))
            .SelectMany(extension => extension.Descendants())
            .Where(element =>
                element.LocalName == "dynamicArrayProperties" &&
                element.NamespaceUri == DynamicArrayNamespace &&
                Boolean(element.GetAttribute("fDynamic", string.Empty).Value) == true &&
                Boolean(element.GetAttribute("fCollapsed", string.Empty).Value) != true)
            .ToArray();
        return matches.Length == 1;
    }

    private static bool? Boolean(string? value) => value?.Trim().ToLowerInvariant() switch
    {
        "1" or "true" => true,
        "0" or "false" or "" => false,
        null => null,
        _ => null,
    };

    private static CodecException Unsupported(CellArtifact cell, string message) =>
        new(
            "unsupported_dynamic_array_edit",
            $"Cell {CellReference(cell.Row, cell.Column)} {message}.",
            CellReference(cell.Row, cell.Column));

    private static string CellReference(uint row, uint column)
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
        return $"{new string(buffer[position..])}{row + 1}";
    }
}
