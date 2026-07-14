using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns the bounded worksheet TableParts -> TableDefinitionPart graph. Cell
// values remain owned by XlsxCodec; query tables, formulas, filters, extensions,
// and other complex table profiles remain opaque and uneditable.
internal sealed class XlsxTableCodec
{
    private static readonly XNamespace Spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly Regex TableName = new("^[A-Za-z_\\\\][A-Za-z0-9_.\\\\]{0,254}$", RegexOptions.CultureInvariant);
    private static readonly Regex R1C1Reference = new("^R[1-9][0-9]*C[1-9][0-9]*$", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
    private static readonly HashSet<string> RootAttributes = new(StringComparer.Ordinal)
    {
        "id", "name", "displayName", "ref", "headerRowCount", "totalsRowShown",
    };
    private readonly WorksheetPart _worksheetPart;
    private readonly string _worksheetPath;
    private readonly List<Entry> _entries = [];

    private sealed class Entry
    {
        internal required TableDefinitionPart Part { get; init; }
        internal required string RelationshipId { get; init; }
        internal required string Path { get; init; }
        internal required XDocument Document { get; init; }
        internal required SpreadsheetTableArtifact Artifact { get; set; }
        internal bool Editable { get; init; }
        internal bool Dirty { get; set; }
    }

    internal XlsxTableCodec(WorksheetPart worksheetPart)
    {
        _worksheetPart = worksheetPart;
        _worksheetPath = worksheetPart.Uri.OriginalString.TrimStart('/');
        var tableParts = worksheetPart.Worksheet?.Elements<TableParts>().ToArray() ?? [];
        if (tableParts.Length > 1) throw Invalid("Worksheet contains more than one tableParts collection.", _worksheetPath);
        foreach (var locator in tableParts.SingleOrDefault()?.Elements<TablePart>() ?? [])
        {
            var relationshipId = locator.Id?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId) || worksheetPart.GetPartById(relationshipId) is not TableDefinitionPart part)
                throw Invalid("Worksheet tableParts contains a dangling or non-table relationship.", _worksheetPath);
            Load(part, relationshipId);
        }
        if (_entries.Select(item => item.Path).Distinct(StringComparer.OrdinalIgnoreCase).Count() != _entries.Count)
            throw Invalid("Worksheet tableParts references one table definition more than once.", _worksheetPath);
    }

    internal IReadOnlyList<SpreadsheetTableArtifact> Read() => _entries.Select(item => item.Artifact.Clone()).ToArray();
    internal IReadOnlySet<string> DirtyPartPaths => _entries.Where(item => item.Dirty).Select(item => item.Path).ToHashSet(StringComparer.OrdinalIgnoreCase);

    internal void Apply(IReadOnlyList<SpreadsheetTableArtifact> desired, bool sourceBound, ref uint nextTableId)
    {
        foreach (var table in desired) Validate(table, _worksheetPath);
        if (!sourceBound)
        {
            if (_entries.Count != 0) throw Invalid("Source-free worksheet unexpectedly contains table definitions.", _worksheetPath);
            foreach (var table in desired) Add(table, ref nextTableId);
            return;
        }

        if (desired.Count != _entries.Count)
            throw Invalid("Source-preserving XLSX export cannot add or remove worksheet tables in this bounded slice.", _worksheetPath);
        var desiredByPath = new Dictionary<string, SpreadsheetTableArtifact>(StringComparer.OrdinalIgnoreCase);
        foreach (var table in desired)
        {
            var path = table.Source?.TablePartPath;
            if (string.IsNullOrWhiteSpace(path) || !desiredByPath.TryAdd(path, table))
                throw Invalid("Source-preserving XLSX export changed or duplicated worksheet table identity.", _worksheetPath);
        }
        foreach (var entry in _entries)
        {
            if (!desiredByPath.TryGetValue(entry.Path, out var table))
                throw Invalid("Source-preserving XLSX export changed worksheet table identity or order.", entry.Path);
            ValidateBinding(table.Source, entry);
            if (!entry.Editable)
            {
                if (HasCompleteSemantics(table)) throw Invalid("The source table uses semantics outside the bounded profile and cannot be replaced losslessly.", entry.Path);
                continue;
            }
            if (!HasCompleteSemantics(table)) throw Invalid("An editable source table is missing its bounded semantic fields.", entry.Path);
            if (SemanticSha256(table).Equals(table.Source.SemanticSha256, StringComparison.OrdinalIgnoreCase)) continue;
            Patch(entry, table);
        }
    }

    internal void Save()
    {
        foreach (var entry in _entries.Where(item => item.Dirty))
        {
            using var stream = entry.Part.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false, OmitXmlDeclaration = false });
            entry.Document.Save(writer);
        }
    }

    internal static void ValidateWorksheet(WorksheetArtifact worksheet)
    {
        if (worksheet.Tables.Count > 65_535) throw Invalid("Worksheet exceeds the bounded table count.", worksheet.Name);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var table in worksheet.Tables)
        {
            Validate(table, worksheet.Name);
            if (HasCompleteSemantics(table) && !names.Add(table.Name)) throw Invalid($"Worksheet contains duplicate table name {table.Name}.", worksheet.Name);
        }
    }

    internal static bool HasCompleteSemantics(SpreadsheetTableArtifact table) =>
        !string.IsNullOrWhiteSpace(table.Name) && !string.IsNullOrWhiteSpace(table.Reference) && table.ColumnNames.Count > 0;

    private void Load(TableDefinitionPart part, string relationshipId)
    {
        byte[] bytes;
        XDocument document;
        using (var source = part.GetStream(FileMode.Open, FileAccess.Read))
        using (var copy = new MemoryStream())
        {
            source.CopyTo(copy);
            bytes = copy.ToArray();
        }
        try
        {
            using var reader = XmlReader.Create(new MemoryStream(bytes, false), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
            document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }
        catch (XmlException exception)
        {
            throw new CodecException("invalid_worksheet_table", "Worksheet table definition is not valid XML.", part.Uri.ToString(), exception);
        }
        var path = part.Uri.OriginalString.TrimStart('/');
        var editable = TryRead(document, out var semantic) && !part.Parts.Any() && !part.ExternalRelationships.Any();
        var binding = new SpreadsheetTableSourceBinding
        {
            WorksheetPartPath = _worksheetPath,
            TablePartPath = path,
            RelationshipId = relationshipId,
            XmlSha256 = Sha256(bytes),
            SemanticSha256 = editable ? SemanticSha256(semantic!) : string.Empty,
            Editable = editable,
        };
        var artifact = semantic ?? new SpreadsheetTableArtifact { Id = $"table/{path}" };
        artifact.Source = binding;
        _entries.Add(new Entry { Part = part, RelationshipId = relationshipId, Path = path, Document = document, Artifact = artifact, Editable = editable });
    }

    private void Add(SpreadsheetTableArtifact table, ref uint nextTableId)
    {
        var relationshipId = NextRelationshipId();
        var part = _worksheetPart.AddNewPart<TableDefinitionPart>(relationshipId);
        var document = Create(table, nextTableId++);
        var entry = new Entry
        {
            Part = part,
            RelationshipId = relationshipId,
            Path = part.Uri.OriginalString.TrimStart('/'),
            Document = document,
            Artifact = table.Clone(),
            Editable = true,
            Dirty = true,
        };
        _entries.Add(entry);
        var tableParts = _worksheetPart.Worksheet!.GetFirstChild<TableParts>();
        if (tableParts is null)
        {
            tableParts = new TableParts();
            var extension = _worksheetPart.Worksheet.GetFirstChild<WorksheetExtensionList>();
            if (extension is null) _worksheetPart.Worksheet.Append(tableParts);
            else _worksheetPart.Worksheet.InsertBefore(tableParts, extension);
        }
        tableParts.Append(new TablePart { Id = relationshipId });
        tableParts.Count = checked((uint)tableParts.ChildElements.Count);
    }

    private string NextRelationshipId()
    {
        var used = _worksheetPart.Parts.Select(item => item.RelationshipId).Concat(_worksheetPart.ExternalRelationships.Select(item => item.Id)).ToHashSet(StringComparer.Ordinal);
        for (var index = 1; ; index++) if (!used.Contains($"rIdTable{index}")) return $"rIdTable{index}";
    }

    private static void Validate(SpreadsheetTableArtifact table, string location)
    {
        if (!HasCompleteSemantics(table))
        {
            if (table.Source is { Editable: false, TablePartPath.Length: > 0, XmlSha256.Length: > 0 }) return;
            throw Invalid("Worksheet table must provide name, reference, and column names.", location);
        }
        if (!TableName.IsMatch(table.Name) || CellReference(table.Name) || R1C1Reference.IsMatch(table.Name) || table.Name.Equals("R", StringComparison.OrdinalIgnoreCase) || table.Name.Equals("C", StringComparison.OrdinalIgnoreCase))
            throw Invalid($"Worksheet table name {table.Name} is invalid.", location);
        var bounds = Range(table.Reference, location);
        var columns = checked((int)(bounds.Right - bounds.Left + 1));
        var rows = checked((int)(bounds.Bottom - bounds.Top + 1));
        if (table.ColumnNames.Count != columns) throw Invalid($"Worksheet table {table.Name} must provide {columns} column names.", location);
        if (rows < (table.HasHeaders ? 1 : 0) + (table.ShowTotals ? 1 : 0) + 1) throw Invalid($"Worksheet table {table.Name} range is too short for its header/totals profile.", location);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var column in table.ColumnNames)
            if (string.IsNullOrWhiteSpace(column) || column.Length > 255 || column.Any(char.IsControl) || !names.Add(column))
                throw Invalid($"Worksheet table {table.Name} has an invalid or duplicate column name.", location);
        if (string.IsNullOrWhiteSpace(table.StyleName) || table.StyleName.Length > 255 || table.StyleName.Any(char.IsControl))
            throw Invalid($"Worksheet table {table.Name} has an invalid style name.", location);
        if (table.ShowFilterButton && !table.HasHeaders) throw Invalid($"Worksheet table {table.Name} cannot show filter buttons without headers.", location);
    }

    private static bool TryRead(XDocument document, out SpreadsheetTableArtifact? artifact)
    {
        artifact = null;
        var root = document.Root;
        if (root?.Name != Spreadsheet + "table" || root.Attributes().Any(item => !item.IsNamespaceDeclaration && (item.Name.Namespace != XNamespace.None || !RootAttributes.Contains(item.Name.LocalName)))) return false;
        if (!uint.TryParse(root.Attribute("id")?.Value, out var tableId) || tableId == 0) return false;
        var name = root.Attribute("name")?.Value;
        var displayName = root.Attribute("displayName")?.Value;
        var reference = root.Attribute("ref")?.Value;
        if (string.IsNullOrWhiteSpace(name) || name != displayName || string.IsNullOrWhiteSpace(reference)) return false;
        if (!TryBool(root.Attribute("headerRowCount")?.Value, defaultValue: true, out var hasHeaders) ||
            !TryBool(root.Attribute("totalsRowShown")?.Value, defaultValue: false, out var showTotals)) return false;
        var children = root.Elements().ToArray();
        if (children.Any(item => item.Name != Spreadsheet + "autoFilter" && item.Name != Spreadsheet + "tableColumns" && item.Name != Spreadsheet + "tableStyleInfo")) return false;
        if (children.Select(item => item.Name).Distinct().Count() != children.Length) return false;
        var expectedIndex = 0;
        if (children.FirstOrDefault()?.Name == Spreadsheet + "autoFilter") expectedIndex++;
        if (children.ElementAtOrDefault(expectedIndex++)?.Name != Spreadsheet + "tableColumns" ||
            children.ElementAtOrDefault(expectedIndex++)?.Name != Spreadsheet + "tableStyleInfo" || expectedIndex != children.Length) return false;
        var columnsRoot = root.Element(Spreadsheet + "tableColumns");
        if (columnsRoot is null || columnsRoot.Attributes().Any(item => !item.IsNamespaceDeclaration && (item.Name.Namespace != XNamespace.None || item.Name.LocalName != "count"))) return false;
        var columns = columnsRoot.Elements().ToArray();
        if (columns.Any(item => item.Name != Spreadsheet + "tableColumn" || item.Elements().Any() || item.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && (attribute.Name.Namespace != XNamespace.None || attribute.Name.LocalName is not ("id" or "name"))))) return false;
        if (!uint.TryParse(columnsRoot.Attribute("count")?.Value, out var columnCount) || columnCount != columns.Length) return false;
        for (var index = 0; index < columns.Length; index++)
            if (!uint.TryParse(columns[index].Attribute("id")?.Value, out var columnId) || columnId != index + 1) return false;
        var names = columns.Select(item => item.Attribute("name")?.Value).ToArray();
        if (names.Any(string.IsNullOrWhiteSpace)) return false;
        var filter = root.Element(Spreadsheet + "autoFilter");
        if (filter is not null && (filter.Elements().Any() || filter.Attributes().Any(item => !item.IsNamespaceDeclaration && (item.Name.Namespace != XNamespace.None || item.Name.LocalName != "ref")) || filter.Attribute("ref")?.Value != reference)) return false;
        var style = root.Element(Spreadsheet + "tableStyleInfo");
        if (style is null || style.Elements().Any() || style.Attributes().Any(item => !item.IsNamespaceDeclaration && (item.Name.Namespace != XNamespace.None || item.Name.LocalName is not ("name" or "showFirstColumn" or "showLastColumn" or "showRowStripes" or "showColumnStripes")))) return false;
        var styleName = style.Attribute("name")?.Value;
        if (string.IsNullOrWhiteSpace(styleName) ||
            !TryBool(style.Attribute("showFirstColumn")?.Value, defaultValue: false, out var showFirstColumn) ||
            !TryBool(style.Attribute("showLastColumn")?.Value, defaultValue: false, out var showLastColumn) ||
            !TryBool(style.Attribute("showRowStripes")?.Value, defaultValue: false, out var showRowStripes) ||
            !TryBool(style.Attribute("showColumnStripes")?.Value, defaultValue: false, out var showColumnStripes)) return false;
        artifact = new SpreadsheetTableArtifact
        {
            Id = $"table/{tableId}",
            Name = name,
            Reference = reference,
            HasHeaders = hasHeaders,
            ShowTotals = showTotals,
            ShowFilterButton = filter is not null,
            StyleName = styleName,
            ShowFirstColumn = showFirstColumn,
            ShowLastColumn = showLastColumn,
            ShowRowStripes = showRowStripes,
            ShowColumnStripes = showColumnStripes,
        };
        artifact.ColumnNames.Add(names!);
        try { Validate(artifact, artifact.Name); }
        catch (CodecException) { artifact = null; return false; }
        return true;
    }

    private static void ValidateBinding(SpreadsheetTableSourceBinding? binding, Entry entry)
    {
        if (binding is null || !binding.WorksheetPartPath.Equals(entry.Artifact.Source.WorksheetPartPath, StringComparison.OrdinalIgnoreCase) ||
            !binding.TablePartPath.Equals(entry.Path, StringComparison.OrdinalIgnoreCase) || binding.RelationshipId != entry.RelationshipId ||
            !binding.XmlSha256.Equals(entry.Artifact.Source.XmlSha256, StringComparison.OrdinalIgnoreCase) || binding.Editable != entry.Editable ||
            !binding.SemanticSha256.Equals(entry.Artifact.Source.SemanticSha256, StringComparison.OrdinalIgnoreCase))
            throw Invalid("Worksheet table source binding does not match the validated source package.", entry.Path);
    }

    private static void Patch(Entry entry, SpreadsheetTableArtifact table)
    {
        var root = entry.Document.Root!;
        root.SetAttributeValue("name", table.Name);
        root.SetAttributeValue("displayName", table.Name);
        root.SetAttributeValue("ref", table.Reference);
        root.SetAttributeValue("headerRowCount", table.HasHeaders ? 1 : 0);
        root.SetAttributeValue("totalsRowShown", table.ShowTotals ? 1 : 0);
        root.Element(Spreadsheet + "autoFilter")?.Remove();
        var columns = root.Element(Spreadsheet + "tableColumns")!;
        var style = root.Element(Spreadsheet + "tableStyleInfo")!;
        if (table.ShowFilterButton) columns.AddBeforeSelf(new XElement(Spreadsheet + "autoFilter", new XAttribute("ref", table.Reference)));
        foreach (var pair in columns.Elements().Zip(table.ColumnNames)) pair.First.SetAttributeValue("name", pair.Second);
        style.SetAttributeValue("name", table.StyleName);
        style.SetAttributeValue("showFirstColumn", table.ShowFirstColumn ? 1 : 0);
        style.SetAttributeValue("showLastColumn", table.ShowLastColumn ? 1 : 0);
        style.SetAttributeValue("showRowStripes", table.ShowRowStripes ? 1 : 0);
        style.SetAttributeValue("showColumnStripes", table.ShowColumnStripes ? 1 : 0);
        entry.Artifact = table.Clone();
        entry.Dirty = true;
    }

    private static XDocument Create(SpreadsheetTableArtifact table, uint tableId)
    {
        var root = new XElement(Spreadsheet + "table",
            new XAttribute(XNamespace.Xmlns + "x", Spreadsheet), new XAttribute("id", tableId), new XAttribute("name", table.Name),
            new XAttribute("displayName", table.Name), new XAttribute("ref", table.Reference),
            new XAttribute("headerRowCount", table.HasHeaders ? 1 : 0), new XAttribute("totalsRowShown", table.ShowTotals ? 1 : 0));
        if (table.ShowFilterButton) root.Add(new XElement(Spreadsheet + "autoFilter", new XAttribute("ref", table.Reference)));
        root.Add(new XElement(Spreadsheet + "tableColumns", new XAttribute("count", table.ColumnNames.Count),
            table.ColumnNames.Select((name, index) => new XElement(Spreadsheet + "tableColumn", new XAttribute("id", index + 1), new XAttribute("name", name)))));
        root.Add(new XElement(Spreadsheet + "tableStyleInfo", new XAttribute("name", table.StyleName),
            new XAttribute("showFirstColumn", table.ShowFirstColumn ? 1 : 0), new XAttribute("showLastColumn", table.ShowLastColumn ? 1 : 0),
            new XAttribute("showRowStripes", table.ShowRowStripes ? 1 : 0), new XAttribute("showColumnStripes", table.ShowColumnStripes ? 1 : 0)));
        return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
    }

    private static string SemanticSha256(SpreadsheetTableArtifact table) => Sha256(Encoding.UTF8.GetBytes(string.Join('\0',
    [
        table.Name, table.Reference, table.HasHeaders.ToString(), table.ShowTotals.ToString(), table.ShowFilterButton.ToString(), table.StyleName,
        table.ShowFirstColumn.ToString(), table.ShowLastColumn.ToString(), table.ShowRowStripes.ToString(), table.ShowColumnStripes.ToString(),
        .. table.ColumnNames,
    ])));

    private static (uint Top, uint Left, uint Bottom, uint Right) Range(string reference, string location)
    {
        var parts = reference.Split(':');
        if (parts.Length is < 1 or > 2 || !TryCell(parts[0], out var first) || !TryCell(parts.Length == 2 ? parts[1] : parts[0], out var second) || first.Row > second.Row || first.Column > second.Column)
            throw Invalid($"Worksheet table reference {reference} is invalid.", location);
        return (first.Row, first.Column, second.Row, second.Column);
    }

    private static bool TryCell(string text, out (uint Row, uint Column) cell)
    {
        cell = default;
        var match = Regex.Match(text, "^\\$?([A-Za-z]{1,3})\\$?([1-9][0-9]*)$", RegexOptions.CultureInvariant);
        if (!match.Success || !uint.TryParse(match.Groups[2].Value, out var row) || row > 1_048_576) return false;
        uint column = 0;
        foreach (var character in match.Groups[1].Value.ToUpperInvariant()) column = checked(column * 26 + (uint)(character - 'A' + 1));
        if (column > 16_384) return false;
        cell = (row - 1, column - 1);
        return true;
    }

    private static bool CellReference(string value) => TryCell(value, out _);
    private static bool TryBool(string? value, bool defaultValue, out bool result)
    {
        if (value is null) { result = defaultValue; return true; }
        if (value is "1" or "true") { result = true; return true; }
        if (value is "0" or "false") { result = false; return true; }
        result = false;
        return false;
    }
    private static string Sha256(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static CodecException Invalid(string message, string? part) => new("invalid_worksheet_table", message, part);
}
