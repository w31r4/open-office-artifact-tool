using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Reads the narrow QueryTablePart profile for one worksheet table. Imported
// external-data graphs are source-bound and read-only: the public projection is
// inspectable, but only the validated source package may supply its XML.
internal sealed class XlsxQueryTableCodec
{
    private enum RefreshProfile { Absent, Recognized, Opaque }
    private enum NestedProfile { Absent, Recognized, Opaque }

    private static readonly XNamespace Spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly HashSet<string> GrowShrinkTypes = new(StringComparer.Ordinal)
    {
        "insertClear", "insertDelete", "overwriteClear",
    };
    private readonly SpreadsheetTableQueryArtifact _sourceArtifact;

    private XlsxQueryTableCodec(
        QueryTablePart part,
        string relationshipId,
        byte[] queryBytes,
        XlsxConnectionCodec connections,
        SpreadsheetTableQueryArtifact artifact)
    {
        RelationshipId = relationshipId;
        Path = part.Uri.OriginalString.TrimStart('/');
        artifact.Source = new SpreadsheetTableQuerySourceBinding
        {
            QueryPartPath = Path,
            RelationshipId = relationshipId,
            XmlSha256 = Sha256(queryBytes),
            SemanticSha256 = SemanticSha256(artifact),
            ConnectionPartPath = connections.Path,
            ConnectionXmlSha256 = connections.PartXmlSha256,
            Editable = false,
        };
        _sourceArtifact = artifact.Clone();
        Artifact = artifact;
    }

    internal string Path { get; }
    internal string RelationshipId { get; }
    internal SpreadsheetTableQueryArtifact Artifact { get; }

    // Returns true when the table owns either no child relationship or exactly
    // one recognized QueryTablePart. False keeps the parent table in its prior
    // opaque/read-only profile without flattening an unknown graph.
    internal static bool TryLoad(TableDefinitionPart tablePart, XlsxConnectionCodec connections, XlsxCellStyleCodec styles, out XlsxQueryTableCodec? codec)
    {
        codec = null;
        if (tablePart.ExternalRelationships.Any()) return false;
        var children = tablePart.Parts.ToArray();
        if (children.Length == 0) return true;
        if (children.Length != 1 || children[0].OpenXmlPart is not QueryTablePart queryPart ||
            queryPart.Parts.Any() || queryPart.ExternalRelationships.Any()) return false;
        if (!TryReadTableIdentity(tablePart, out var tableColumnIds, out var tableBounds) ||
            !TryReadPart(queryPart, out var queryBytes, out var queryDocument) ||
            !TryReadQuery(queryDocument!, tableColumnIds, styles, tableBounds, out var artifact, out _, out _,
                out _, out _, out _, out _) ||
            !connections.Contains(artifact!.ConnectionId)) return false;

        codec = new XlsxQueryTableCodec(
            queryPart,
            children[0].RelationshipId,
            queryBytes!,
            connections,
            artifact!);
        return true;
    }

    internal static void Validate(SpreadsheetTableQueryArtifact? query, string location)
    {
        if (query is null) return;
        if (query.Source is null)
            throw Unsupported("Source-free XLSX authoring cannot fabricate a QueryTable/external-connection graph.", location);
    }

    internal void Apply(SpreadsheetTableQueryArtifact? desired, bool sourceBound)
    {
        if (!sourceBound)
            throw Unsupported("Source-free XLSX authoring cannot fabricate a QueryTable/external-connection graph.", Path);
        if (desired is null)
            throw Unsupported("Imported worksheet QueryTables are read-only and cannot be removed.", Path);
        ValidateBinding(desired.Source);
        if (SemanticSha256(desired).Equals(_sourceArtifact.Source.SemanticSha256, StringComparison.OrdinalIgnoreCase)) return;
        throw Unsupported("Imported worksheet QueryTables are read-only and cannot be edited.", Path);
    }

    private void ValidateBinding(SpreadsheetTableQuerySourceBinding? binding)
    {
        var source = _sourceArtifact.Source;
        if (binding is null || !binding.QueryPartPath.Equals(Path, StringComparison.OrdinalIgnoreCase) ||
            binding.RelationshipId != RelationshipId ||
            !binding.XmlSha256.Equals(source.XmlSha256, StringComparison.OrdinalIgnoreCase) ||
            !binding.SemanticSha256.Equals(source.SemanticSha256, StringComparison.OrdinalIgnoreCase) ||
            !binding.ConnectionPartPath.Equals(source.ConnectionPartPath, StringComparison.OrdinalIgnoreCase) ||
            !binding.ConnectionXmlSha256.Equals(source.ConnectionXmlSha256, StringComparison.OrdinalIgnoreCase) ||
            binding.Editable != source.Editable)
            throw Unsupported("Worksheet QueryTable source binding does not match the validated source package.", Path);
    }

    private static void ValidateRefresh(SpreadsheetTableQueryRefreshArtifact refresh, string location)
    {
        if (refresh.Fields.Count > 16_384 || refresh.DeletedFieldNames.Count > 16_384 ||
            refresh.HasMinimumVersion && refresh.MinimumVersion > byte.MaxValue ||
            refresh.HasUnboundColumnsLeft && refresh.UnboundColumnsLeft > 16_384 ||
            refresh.HasUnboundColumnsRight && refresh.UnboundColumnsRight > 16_384)
            throw Invalid("Worksheet query refresh metadata exceeds the bounded profile.", location);
        var ids = new HashSet<uint>();
        var tableColumnIds = new HashSet<uint>();
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var rowNumberFields = 0;
        foreach (var field in refresh.Fields)
        {
            if (field.Id == 0 || !ids.Add(field.Id) ||
                field.HasName && (string.IsNullOrWhiteSpace(field.Name) || field.Name.Length > 255 || field.Name.Any(char.IsControl) || !names.Add(field.Name)) ||
                field.HasTableColumnId && (field.TableColumnId == 0 || !tableColumnIds.Add(field.TableColumnId)))
                throw Invalid("Worksheet query refresh has an invalid or duplicate field profile.", location);
            if (field.HasRowNumbers && field.RowNumbers && ++rowNumberFields > 1)
                throw Invalid("Worksheet query refresh cannot identify more than one row-number field.", location);
            if (field.HasFillFormulas && field.FillFormulas && (!field.HasDataBound || field.DataBound))
                throw Invalid("Worksheet query refresh formula fields must be explicitly unbound from external data.", location);
            if (field.HasClipped && field.Clipped && (!field.HasDataBound || !field.DataBound))
                throw Invalid("Worksheet query refresh clipped fields must be explicitly bound to external data.", location);
        }
        var deletedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var name in refresh.DeletedFieldNames)
            if (string.IsNullOrWhiteSpace(name) || name.Length > 255 || name.Any(char.IsControl) || !deletedNames.Add(name))
                throw Invalid("Worksheet query refresh has an invalid or duplicate deleted-field name.", location);
        if (refresh.HasNextId && (refresh.NextId == 0 || ids.Contains(refresh.NextId)))
            throw Invalid("Worksheet query refresh next_id must identify an unused positive field ID.", location);
    }

    private static bool TryReadQuery(
        XDocument document,
        IReadOnlySet<uint> tableColumnIds,
        XlsxCellStyleCodec styles,
        (uint Top, uint Left, uint Bottom, uint Right) tableBounds,
        out SpreadsheetTableQueryArtifact? artifact,
        out RefreshProfile refreshProfile,
        out XElement? refreshElement,
        out NestedProfile deletedFieldsProfile,
        out XElement? deletedFieldsElement,
        out NestedProfile sortProfile,
        out XlsxQuerySortStateCodec? sort)
    {
        artifact = null;
        refreshProfile = RefreshProfile.Absent;
        refreshElement = null;
        deletedFieldsProfile = NestedProfile.Absent;
        deletedFieldsElement = null;
        sortProfile = NestedProfile.Absent;
        sort = null;
        var root = document.Root;
        if (root?.Name != Spreadsheet + "queryTable") return false;
        var name = root.Attribute("name")?.Value;
        if (string.IsNullOrWhiteSpace(name) || name.Length > 255 || name.Any(char.IsControl) ||
            !uint.TryParse(root.Attribute("connectionId")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var connectionId) || connectionId == 0)
            return false;
        var result = new SpreadsheetTableQueryArtifact { Name = name, ConnectionId = connectionId };
        if (!ReadOptionalBool(root, "headers", value => result.Headers = value) ||
            !ReadOptionalBool(root, "rowNumbers", value => result.RowNumbers = value) ||
            !ReadOptionalBool(root, "disableRefresh", value => result.DisableRefresh = value) ||
            !ReadOptionalBool(root, "backgroundRefresh", value => result.BackgroundRefresh = value) ||
            !ReadOptionalBool(root, "firstBackgroundRefresh", value => result.FirstBackgroundRefresh = value) ||
            !ReadOptionalBool(root, "refreshOnLoad", value => result.RefreshOnLoad = value) ||
            !ReadOptionalBool(root, "fillFormulas", value => result.FillFormulas = value) ||
            !ReadOptionalBool(root, "removeDataOnSave", value => result.RemoveDataOnSave = value) ||
            !ReadOptionalBool(root, "disableEdit", value => result.DisableEdit = value) ||
            !ReadOptionalBool(root, "preserveFormatting", value => result.PreserveFormatting = value) ||
            !ReadOptionalBool(root, "adjustColumnWidth", value => result.AdjustColumnWidth = value) ||
            !ReadOptionalBool(root, "intermediate", value => result.Intermediate = value) ||
            !ReadOptionalBool(root, "applyNumberFormats", value => result.ApplyNumberFormats = value) ||
            !ReadOptionalBool(root, "applyBorderFormats", value => result.ApplyBorderFormats = value) ||
            !ReadOptionalBool(root, "applyFontFormats", value => result.ApplyFontFormats = value) ||
            !ReadOptionalBool(root, "applyPatternFormats", value => result.ApplyPatternFormats = value) ||
            !ReadOptionalBool(root, "applyAlignmentFormats", value => result.ApplyAlignmentFormats = value) ||
            !ReadOptionalBool(root, "applyWidthHeightFormats", value => result.ApplyWidthHeightFormats = value)) return false;
        if (root.Attribute("growShrinkType") is { } growShrink)
        {
            if (!GrowShrinkTypes.Contains(growShrink.Value)) return false;
            result.GrowShrinkType = growShrink.Value;
        }
        if (root.Attribute("autoFormatId") is { } autoFormat)
        {
            if (!uint.TryParse(autoFormat.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var autoFormatId)) return false;
            result.AutoFormatId = autoFormatId;
        }
        var refreshElements = root.Elements(Spreadsheet + "queryTableRefresh").ToArray();
        if (refreshElements.Length == 1)
        {
            refreshElement = refreshElements[0];
            if (TryReadRefresh(refreshElement, tableColumnIds, styles, tableBounds, out var refresh,
                out deletedFieldsProfile, out deletedFieldsElement, out sortProfile, out sort))
            {
                result.Refresh = refresh;
                refreshProfile = RefreshProfile.Recognized;
            }
            else refreshProfile = RefreshProfile.Opaque;
        }
        else if (refreshElements.Length > 1) refreshProfile = RefreshProfile.Opaque;
        artifact = result;
        return true;
    }

    private static bool TryReadTableIdentity(
        TableDefinitionPart part,
        out HashSet<uint> ids,
        out (uint Top, uint Left, uint Bottom, uint Right) bounds)
    {
        ids = [];
        bounds = default;
        var columns = part.Table?.TableColumns?.Elements<TableColumn>().ToArray();
        if (columns is null || columns.Length == 0 || part.Table?.Reference?.Value is not string reference ||
            !XlsxQuerySortStateCodec.TryRange(reference, out bounds)) return false;
        foreach (var column in columns)
        {
            var id = column.Id?.Value;
            if (id is null or 0 || !ids.Add(id.Value)) return false;
        }
        return true;
    }

    private static bool TryReadRefresh(
        XElement element,
        IReadOnlySet<uint> tableColumnIds,
        XlsxCellStyleCodec styles,
        (uint Top, uint Left, uint Bottom, uint Right) tableBounds,
        out SpreadsheetTableQueryRefreshArtifact refresh,
        out NestedProfile deletedFieldsProfile,
        out XElement? deletedFieldsElement,
        out NestedProfile sortProfile,
        out XlsxQuerySortStateCodec? sort)
    {
        var result = new SpreadsheetTableQueryRefreshArtifact();
        refresh = result;
        deletedFieldsProfile = NestedProfile.Absent;
        deletedFieldsElement = null;
        sortProfile = NestedProfile.Absent;
        sort = null;
        if (!ReadOptionalBool(element, "preserveSortFilterLayout", value => result.PreserveSortFilterLayout = value) ||
            !ReadOptionalBool(element, "fieldIdWrapped", value => result.FieldIdWrapped = value) ||
            !ReadOptionalBool(element, "headersInLastRefresh", value => result.HeadersInLastRefresh = value) ||
            !ReadOptionalUInt(element, "minimumVersion", value => result.MinimumVersion = value) ||
            !ReadOptionalUInt(element, "nextId", value => result.NextId = value) ||
            !ReadOptionalUInt(element, "unboundColumnsLeft", value => result.UnboundColumnsLeft = value) ||
            !ReadOptionalUInt(element, "unboundColumnsRight", value => result.UnboundColumnsRight = value)) return false;

        var fieldCollections = element.Elements(Spreadsheet + "queryTableFields").ToArray();
        if (fieldCollections.Length > 1) return false;
        if (fieldCollections.Length == 1)
        {
            var fieldElements = fieldCollections[0].Elements().ToArray();
            if (fieldElements.Any(child => child.Name != Spreadsheet + "queryTableField") ||
                !uint.TryParse(fieldCollections[0].Attribute("count")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var count) ||
                count != fieldElements.Length || count > 16_384) return false;
            foreach (var fieldElement in fieldElements)
            {
                if (fieldElement.Elements().Any(child => child.Name != Spreadsheet + "extLst") ||
                    fieldElement.Elements(Spreadsheet + "extLst").Skip(1).Any() ||
                    !uint.TryParse(fieldElement.Attribute("id")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var id) || id == 0)
                    return false;
                var field = new SpreadsheetTableQueryFieldArtifact { Id = id };
                if (fieldElement.Attribute("name") is { } name) field.Name = name.Value;
                if (!ReadOptionalBool(fieldElement, "dataBound", value => field.DataBound = value) ||
                    !ReadOptionalBool(fieldElement, "rowNumbers", value => field.RowNumbers = value) ||
                    !ReadOptionalBool(fieldElement, "fillFormulas", value => field.FillFormulas = value) ||
                    !ReadOptionalBool(fieldElement, "clipped", value => field.Clipped = value)) return false;
                if (fieldElement.Attribute("tableColumnId") is { } tableColumn)
                {
                    if (!uint.TryParse(tableColumn.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var tableColumnId) ||
                        tableColumnId == 0 || !tableColumnIds.Contains(tableColumnId)) return false;
                    field.TableColumnId = tableColumnId;
                }
                result.Fields.Add(field);
            }
        }
        var deletedCollections = element.Elements(Spreadsheet + "queryTableDeletedFields").ToArray();
        if (deletedCollections.Length == 1)
        {
            deletedFieldsElement = deletedCollections[0];
            if (TryReadDeletedFields(deletedFieldsElement, out var names))
            {
                result.DeletedFieldNames.Add(names);
                deletedFieldsProfile = NestedProfile.Recognized;
            }
            else deletedFieldsProfile = NestedProfile.Opaque;
        }
        else if (deletedCollections.Length > 1) deletedFieldsProfile = NestedProfile.Opaque;
        var sortElements = element.Elements(Spreadsheet + "sortState").ToArray();
        if (sortElements.Length == 1)
        {
            if (XlsxQuerySortStateCodec.TryCreate(sortElements[0], styles, tableBounds, out sort))
            {
                result.SortState = sort!.Artifact.Clone();
                sortProfile = NestedProfile.Recognized;
            }
            else sortProfile = NestedProfile.Opaque;
        }
        else if (sortElements.Length > 1) sortProfile = NestedProfile.Opaque;
        try
        {
            ValidateRefresh(result, string.Empty);
            refresh = result;
            return true;
        }
        catch (CodecException)
        {
            refresh = new SpreadsheetTableQueryRefreshArtifact();
            return false;
        }
    }

    private static bool TryReadDeletedFields(XElement element, out IReadOnlyList<string> names)
    {
        names = [];
        var fields = element.Elements().ToArray();
        if (element.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration &&
                (attribute.Name.Namespace != XNamespace.None || attribute.Name.LocalName != "count")) ||
            fields.Any(child => child.Name != Spreadsheet + "deletedField") ||
            !uint.TryParse(element.Attribute("count")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var count) ||
            count != fields.Length || count > 16_384) return false;
        var result = new List<string>(fields.Length);
        var unique = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var field in fields)
        {
            if (field.Elements().Any() || field.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration &&
                (attribute.Name.Namespace != XNamespace.None || attribute.Name.LocalName != "name")) ||
                field.Attribute("name")?.Value is not string name || string.IsNullOrWhiteSpace(name) || name.Length > 255 ||
                name.Any(char.IsControl) || !unique.Add(name)) return false;
            result.Add(name);
        }
        names = result;
        return true;
    }

    private static bool TryReadPart(OpenXmlPart part, out byte[]? bytes, out XDocument? document)
    {
        bytes = null;
        document = null;
        try
        {
            using var source = part.GetStream(FileMode.Open, FileAccess.Read);
            using var copy = new MemoryStream();
            source.CopyTo(copy);
            bytes = copy.ToArray();
            using var reader = XmlReader.Create(new MemoryStream(bytes, writable: false), new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
            });
            document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
            return true;
        }
        catch (XmlException)
        {
            return false;
        }
    }

    private static bool ReadOptionalBool(XElement root, string name, Action<bool> set)
    {
        if (root.Attribute(name) is not { } attribute) return true;
        if (!TryBool(attribute.Value, out var value)) return false;
        set(value);
        return true;
    }

    private static bool ReadOptionalUInt(XElement root, string name, Action<uint> set)
    {
        if (root.Attribute(name) is not { } attribute) return true;
        if (!uint.TryParse(attribute.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var value)) return false;
        set(value);
        return true;
    }

    private static bool TryBool(string value, out bool parsed)
    {
        if (value is "1" or "true") { parsed = true; return true; }
        if (value is "0" or "false") { parsed = false; return true; }
        parsed = false;
        return false;
    }

    private static string SemanticSha256(SpreadsheetTableQueryArtifact query)
    {
        var values = new List<string>
        {
            query.Name,
            query.ConnectionId.ToString(CultureInfo.InvariantCulture),
            Optional(query.HasHeaders, query.Headers),
            Optional(query.HasRowNumbers, query.RowNumbers),
            Optional(query.HasDisableRefresh, query.DisableRefresh),
            Optional(query.HasBackgroundRefresh, query.BackgroundRefresh),
            Optional(query.HasFirstBackgroundRefresh, query.FirstBackgroundRefresh),
            Optional(query.HasRefreshOnLoad, query.RefreshOnLoad),
            query.HasGrowShrinkType ? query.GrowShrinkType : "<absent>",
            Optional(query.HasFillFormulas, query.FillFormulas),
            Optional(query.HasRemoveDataOnSave, query.RemoveDataOnSave),
            Optional(query.HasDisableEdit, query.DisableEdit),
            Optional(query.HasPreserveFormatting, query.PreserveFormatting),
            Optional(query.HasAdjustColumnWidth, query.AdjustColumnWidth),
            Optional(query.HasIntermediate, query.Intermediate),
            Optional(query.HasAutoFormatId, query.AutoFormatId),
            Optional(query.HasApplyNumberFormats, query.ApplyNumberFormats),
            Optional(query.HasApplyBorderFormats, query.ApplyBorderFormats),
            Optional(query.HasApplyFontFormats, query.ApplyFontFormats),
            Optional(query.HasApplyPatternFormats, query.ApplyPatternFormats),
            Optional(query.HasApplyAlignmentFormats, query.ApplyAlignmentFormats),
            Optional(query.HasApplyWidthHeightFormats, query.ApplyWidthHeightFormats),
        };
        if (query.Refresh is null) values.Add("<refresh-absent>");
        else
        {
            var refresh = query.Refresh;
            values.Add("<refresh>");
            values.Add(Optional(refresh.HasPreserveSortFilterLayout, refresh.PreserveSortFilterLayout));
            values.Add(Optional(refresh.HasFieldIdWrapped, refresh.FieldIdWrapped));
            values.Add(Optional(refresh.HasHeadersInLastRefresh, refresh.HeadersInLastRefresh));
            values.Add(Optional(refresh.HasMinimumVersion, refresh.MinimumVersion));
            values.Add(Optional(refresh.HasNextId, refresh.NextId));
            values.Add(Optional(refresh.HasUnboundColumnsLeft, refresh.UnboundColumnsLeft));
            values.Add(Optional(refresh.HasUnboundColumnsRight, refresh.UnboundColumnsRight));
            foreach (var field in refresh.Fields)
            {
                values.Add("<field>");
                values.Add(field.Id.ToString(CultureInfo.InvariantCulture));
                values.Add(field.HasName ? field.Name : "<absent>");
                values.Add(Optional(field.HasDataBound, field.DataBound));
                values.Add(Optional(field.HasRowNumbers, field.RowNumbers));
                values.Add(Optional(field.HasFillFormulas, field.FillFormulas));
                values.Add(Optional(field.HasClipped, field.Clipped));
                values.Add(Optional(field.HasTableColumnId, field.TableColumnId));
            }
            values.Add("<deleted-fields>");
            values.AddRange(refresh.DeletedFieldNames);
            values.AddRange(XlsxQuerySortStateCodec.Semantics(refresh.SortState));
        }
        return Sha256(Encoding.UTF8.GetBytes(string.Join('\0', values)));
    }

    private static string Optional(bool hasValue, bool value) => hasValue ? value ? "1" : "0" : "<absent>";
    private static string Optional(bool hasValue, uint value) => hasValue ? value.ToString(CultureInfo.InvariantCulture) : "<absent>";
    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static CodecException Invalid(string message, string? location = null) => new("invalid_worksheet_table", message, location);
    private static CodecException Unsupported(string message, string? location = null) => new("unsupported_query_table_edit", message, location);
}
