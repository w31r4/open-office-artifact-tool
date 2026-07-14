using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns the narrow, source-bound QueryTablePart decision for one worksheet
// table. The public model sees root query policy and a stable connection ID;
// the source package continues to own connection definitions, refresh history,
// query fields, extensions, and all other XML. This module never authors a new
// external-data graph.
internal sealed class XlsxQueryTableCodec
{
    private static readonly XNamespace Spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly HashSet<string> GrowShrinkTypes = new(StringComparer.Ordinal)
    {
        "insertClear", "insertDelete", "overwriteClear",
    };
    private readonly QueryTablePart _part;
    private readonly XDocument _document;
    private readonly HashSet<uint> _connectionIds;
    private readonly SpreadsheetTableQueryArtifact _sourceArtifact;

    private XlsxQueryTableCodec(
        QueryTablePart part,
        string relationshipId,
        XDocument document,
        byte[] queryBytes,
        string connectionPartPath,
        byte[] connectionBytes,
        HashSet<uint> connectionIds,
        SpreadsheetTableQueryArtifact artifact)
    {
        _part = part;
        RelationshipId = relationshipId;
        _document = document;
        _connectionIds = connectionIds;
        Path = part.Uri.OriginalString.TrimStart('/');
        artifact.Source = new SpreadsheetTableQuerySourceBinding
        {
            QueryPartPath = Path,
            RelationshipId = relationshipId,
            XmlSha256 = Sha256(queryBytes),
            SemanticSha256 = SemanticSha256(artifact),
            ConnectionPartPath = connectionPartPath,
            ConnectionXmlSha256 = Sha256(connectionBytes),
            Editable = true,
        };
        _sourceArtifact = artifact.Clone();
        Artifact = artifact;
    }

    internal string Path { get; }
    internal string RelationshipId { get; }
    internal SpreadsheetTableQueryArtifact Artifact { get; private set; }
    internal bool Dirty { get; private set; }

    // Returns true when the table owns either no child relationship or exactly
    // one recognized QueryTablePart. False keeps the parent table in its prior
    // opaque/read-only profile without flattening an unknown graph.
    internal static bool TryLoad(TableDefinitionPart tablePart, WorkbookPart workbookPart, out XlsxQueryTableCodec? codec)
    {
        codec = null;
        if (tablePart.ExternalRelationships.Any()) return false;
        var children = tablePart.Parts.ToArray();
        if (children.Length == 0) return true;
        if (children.Length != 1 || children[0].OpenXmlPart is not QueryTablePart queryPart ||
            queryPart.Parts.Any() || queryPart.ExternalRelationships.Any()) return false;
        var connectionsPart = workbookPart.ConnectionsPart;
        if (connectionsPart is null) return false;

        if (!TryReadPart(queryPart, out var queryBytes, out var queryDocument) ||
            !TryReadPart(connectionsPart, out var connectionBytes, out var connectionDocument) ||
            !TryReadConnections(connectionDocument!, out var connectionIds) ||
            !TryReadQuery(queryDocument!, out var artifact) ||
            !connectionIds.Contains(artifact!.ConnectionId)) return false;

        codec = new XlsxQueryTableCodec(
            queryPart,
            children[0].RelationshipId,
            queryDocument!,
            queryBytes!,
            connectionsPart.Uri.OriginalString.TrimStart('/'),
            connectionBytes!,
            connectionIds,
            artifact!);
        return true;
    }

    internal static void Validate(SpreadsheetTableQueryArtifact? query, string location)
    {
        if (query is null) return;
        if (string.IsNullOrWhiteSpace(query.Name) || query.Name.Length > 255 || query.Name.Any(char.IsControl))
            throw Invalid("Worksheet query table has an invalid name.", location);
        if (query.ConnectionId == 0)
            throw Invalid("Worksheet query table connection_id must be positive.", location);
        if (query.HasGrowShrinkType && !GrowShrinkTypes.Contains(query.GrowShrinkType))
            throw Invalid($"Worksheet query table grow/shrink type {query.GrowShrinkType} is invalid.", location);
    }

    internal void Apply(SpreadsheetTableQueryArtifact? desired, bool sourceBound)
    {
        if (!sourceBound)
            throw Invalid("Source-free XLSX authoring cannot fabricate a QueryTable/external-connection graph.", Path);
        if (desired is null)
            throw Invalid("Source-preserving XLSX export cannot remove an imported QueryTable graph in this bounded slice.", Path);
        Validate(desired, Path);
        ValidateBinding(desired.Source);
        if (!_connectionIds.Contains(desired.ConnectionId))
            throw Invalid($"Worksheet query table connection_id {desired.ConnectionId} does not identify a connection in the validated source ConnectionsPart.", Path);
        if (SemanticSha256(desired).Equals(_sourceArtifact.Source.SemanticSha256, StringComparison.OrdinalIgnoreCase)) return;
        Patch(desired);
    }

    internal void Save()
    {
        if (!Dirty) return;
        using var stream = _part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = XmlWriter.Create(stream, new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(false),
            Indent = false,
            OmitXmlDeclaration = false,
        });
        _document.Save(writer);
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
            throw Invalid("Worksheet query table source binding does not match the validated source package.", Path);
    }

    private void Patch(SpreadsheetTableQueryArtifact query)
    {
        var root = _document.Root!;
        root.SetAttributeValue("name", query.Name);
        root.SetAttributeValue("connectionId", query.ConnectionId.ToString(CultureInfo.InvariantCulture));
        SetOptional(root, "headers", query.HasHeaders, query.Headers);
        SetOptional(root, "rowNumbers", query.HasRowNumbers, query.RowNumbers);
        SetOptional(root, "disableRefresh", query.HasDisableRefresh, query.DisableRefresh);
        SetOptional(root, "backgroundRefresh", query.HasBackgroundRefresh, query.BackgroundRefresh);
        SetOptional(root, "firstBackgroundRefresh", query.HasFirstBackgroundRefresh, query.FirstBackgroundRefresh);
        SetOptional(root, "refreshOnLoad", query.HasRefreshOnLoad, query.RefreshOnLoad);
        root.SetAttributeValue("growShrinkType", query.HasGrowShrinkType ? query.GrowShrinkType : null);
        SetOptional(root, "fillFormulas", query.HasFillFormulas, query.FillFormulas);
        SetOptional(root, "removeDataOnSave", query.HasRemoveDataOnSave, query.RemoveDataOnSave);
        SetOptional(root, "disableEdit", query.HasDisableEdit, query.DisableEdit);
        SetOptional(root, "preserveFormatting", query.HasPreserveFormatting, query.PreserveFormatting);
        SetOptional(root, "adjustColumnWidth", query.HasAdjustColumnWidth, query.AdjustColumnWidth);
        SetOptional(root, "intermediate", query.HasIntermediate, query.Intermediate);
        root.SetAttributeValue("autoFormatId", query.HasAutoFormatId ? query.AutoFormatId.ToString(CultureInfo.InvariantCulture) : null);
        SetOptional(root, "applyNumberFormats", query.HasApplyNumberFormats, query.ApplyNumberFormats);
        SetOptional(root, "applyBorderFormats", query.HasApplyBorderFormats, query.ApplyBorderFormats);
        SetOptional(root, "applyFontFormats", query.HasApplyFontFormats, query.ApplyFontFormats);
        SetOptional(root, "applyPatternFormats", query.HasApplyPatternFormats, query.ApplyPatternFormats);
        SetOptional(root, "applyAlignmentFormats", query.HasApplyAlignmentFormats, query.ApplyAlignmentFormats);
        SetOptional(root, "applyWidthHeightFormats", query.HasApplyWidthHeightFormats, query.ApplyWidthHeightFormats);
        Artifact = query.Clone();
        Dirty = true;
    }

    private static bool TryReadQuery(XDocument document, out SpreadsheetTableQueryArtifact? artifact)
    {
        artifact = null;
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
        artifact = result;
        return true;
    }

    private static bool TryReadConnections(XDocument document, out HashSet<uint> connectionIds)
    {
        connectionIds = [];
        var root = document.Root;
        if (root?.Name != Spreadsheet + "connections") return false;
        foreach (var connection in root.Elements(Spreadsheet + "connection"))
        {
            if (!uint.TryParse(connection.Attribute("id")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var id) || id == 0 || !connectionIds.Add(id))
                return false;
        }
        return connectionIds.Count > 0;
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

    private static bool TryBool(string value, out bool parsed)
    {
        if (value is "1" or "true") { parsed = true; return true; }
        if (value is "0" or "false") { parsed = false; return true; }
        parsed = false;
        return false;
    }

    private static void SetOptional(XElement root, string name, bool hasValue, bool value) =>
        root.SetAttributeValue(name, hasValue ? value ? "1" : "0" : null);

    private static string SemanticSha256(SpreadsheetTableQueryArtifact query) => Sha256(Encoding.UTF8.GetBytes(string.Join('\0',
    [
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
        query.HasAutoFormatId ? query.AutoFormatId.ToString(CultureInfo.InvariantCulture) : "<absent>",
        Optional(query.HasApplyNumberFormats, query.ApplyNumberFormats),
        Optional(query.HasApplyBorderFormats, query.ApplyBorderFormats),
        Optional(query.HasApplyFontFormats, query.ApplyFontFormats),
        Optional(query.HasApplyPatternFormats, query.ApplyPatternFormats),
        Optional(query.HasApplyAlignmentFormats, query.ApplyAlignmentFormats),
        Optional(query.HasApplyWidthHeightFormats, query.ApplyWidthHeightFormats),
    ])));

    private static string Optional(bool hasValue, bool value) => hasValue ? value ? "1" : "0" : "<absent>";
    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static CodecException Invalid(string message, string? location = null) => new("invalid_worksheet_table", message, location);
}
