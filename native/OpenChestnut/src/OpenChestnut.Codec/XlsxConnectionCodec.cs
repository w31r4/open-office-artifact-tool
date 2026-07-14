using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns the workbook-level ConnectionsPart exactly once. The bounded public
// projection contains only non-secret root metadata and refresh policy for the
// first bounded database/type-5 profile;
// provider strings, commands, credentials, paths, child graphs, extensions,
// and unrecognized root attributes remain in the source XML.
internal sealed class XlsxConnectionCodec
{
    private const uint MaxConnections = 4_096;
    private const uint MaxIntervalMinutes = 32_767;
    private static readonly XNamespace Spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private readonly ConnectionsPart? _part;
    private readonly XDocument? _document;
    private readonly List<Entry> _entries = [];
    private readonly HashSet<uint> _connectionIds = [];

    private sealed class Entry
    {
        internal required XElement Element { get; init; }
        internal required uint ConnectionId { get; init; }
        internal SpreadsheetConnectionArtifact? SourceArtifact { get; init; }
    }

    internal XlsxConnectionCodec(WorkbookPart workbookPart)
    {
        _part = workbookPart.ConnectionsPart;
        if (_part is null)
        {
            IsReadable = true;
            return;
        }

        Path = _part.Uri.OriginalString.TrimStart('/');
        if (!TryReadPart(_part, out var bytes, out var document) || document!.Root?.Name != Spreadsheet + "connections") return;
        _document = document;
        PartXmlSha256 = Sha256(bytes!);
        var children = document.Root.Elements().ToArray();
        if (children.Length > MaxConnections || children.Any(child => child.Name != Spreadsheet + "connection")) return;
        foreach (var element in children)
        {
            if (!uint.TryParse(element.Attribute("id")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var id) ||
                id == 0 || !_connectionIds.Add(id))
            {
                _connectionIds.Clear();
                _entries.Clear();
                return;
            }

            SpreadsheetConnectionArtifact? artifact = null;
            if (TryReadConnection(element, out var recognized))
            {
                recognized!.Source = new SpreadsheetConnectionSourceBinding
                {
                    PartPath = Path,
                    PartXmlSha256 = PartXmlSha256,
                    ConnectionXmlSha256 = ElementSha256(element),
                    SemanticSha256 = SemanticSha256(recognized),
                    Editable = true,
                };
                artifact = recognized;
            }
            _entries.Add(new Entry { Element = element, ConnectionId = id, SourceArtifact = artifact });
        }
        IsReadable = true;
    }

    internal bool IsReadable { get; }
    internal string Path { get; } = string.Empty;
    internal string PartXmlSha256 { get; } = string.Empty;
    internal bool Dirty { get; private set; }
    internal bool Contains(uint connectionId) => IsReadable && _connectionIds.Contains(connectionId);
    internal IReadOnlyList<SpreadsheetConnectionArtifact> Read() => _entries
        .Where(entry => entry.SourceArtifact is not null)
        .Select(entry => entry.SourceArtifact!.Clone())
        .ToArray();

    internal void Apply(IReadOnlyList<SpreadsheetConnectionArtifact> desired, bool sourceBound)
    {
        if (!sourceBound)
        {
            if (desired.Count > 0)
                throw Invalid("Source-free XLSX authoring cannot fabricate workbook connection definitions.");
            return;
        }
        if (!IsReadable)
        {
            if (desired.Count > 0)
                throw Invalid("Source-preserving XLSX export cannot replace an opaque ConnectionsPart.", Path);
            return;
        }

        var recognized = _entries.Where(entry => entry.SourceArtifact is not null).ToArray();
        if (desired.Count != recognized.Length)
            throw Invalid("Source-preserving XLSX export cannot add or remove recognized workbook connections.", Path);
        for (var index = 0; index < recognized.Length; index++)
        {
            var entry = recognized[index];
            var source = entry.SourceArtifact!;
            var target = desired[index];
            Validate(target, Path);
            ValidateBinding(target.Source, source.Source, entry.ConnectionId);
            if (target.ConnectionId != source.ConnectionId || target.Type != source.Type ||
                target.RefreshedVersion != source.RefreshedVersion)
                throw Invalid("Source-preserving XLSX export cannot reorder or change connection id/type/version identity.", Path);
            if (SemanticSha256(target).Equals(source.Source.SemanticSha256, StringComparison.OrdinalIgnoreCase)) continue;
            Patch(entry.Element, target);
            Dirty = true;
        }
    }

    internal void Save()
    {
        if (!Dirty || _part is null || _document is null) return;
        using var stream = _part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = XmlWriter.Create(stream, new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(false),
            Indent = false,
            OmitXmlDeclaration = false,
        });
        _document.Save(writer);
    }

    private static bool TryReadConnection(XElement element, out SpreadsheetConnectionArtifact? artifact)
    {
        artifact = null;
        var name = element.Attribute("name")?.Value;
        if (!uint.TryParse(element.Attribute("id")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var id) || id == 0 ||
            string.IsNullOrWhiteSpace(name) || name.Length > 255 || name.Any(char.IsControl) ||
            !uint.TryParse(element.Attribute("type")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var type) || type != 5 ||
            !uint.TryParse(element.Attribute("refreshedVersion")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var refreshedVersion) ||
            refreshedVersion > byte.MaxValue)
            return false;

        var result = new SpreadsheetConnectionArtifact
        {
            ConnectionId = id,
            Name = name,
            Type = type,
            RefreshedVersion = refreshedVersion,
        };
        if (element.Attribute("description") is { } description) result.Description = description.Value;
        if (!ReadOptionalBool(element, "keepAlive", value => result.KeepAlive = value) ||
            !ReadOptionalUInt(element, "interval", value => result.IntervalMinutes = value) ||
            !ReadOptionalBool(element, "background", value => result.Background = value) ||
            !ReadOptionalBool(element, "refreshOnLoad", value => result.RefreshOnLoad = value) ||
            !ReadOptionalBool(element, "saveData", value => result.SaveData = value)) return false;
        try
        {
            Validate(result, string.Empty);
            artifact = result;
            return true;
        }
        catch (CodecException)
        {
            return false;
        }
    }

    private static void Validate(SpreadsheetConnectionArtifact connection, string location)
    {
        if (connection.ConnectionId == 0 || connection.Type != 5 || connection.RefreshedVersion > byte.MaxValue ||
            string.IsNullOrWhiteSpace(connection.Name) || connection.Name.Length > 255 || connection.Name.Any(char.IsControl) ||
            connection.HasDescription && (connection.Description.Length > 32_767 || connection.Description.Any(char.IsControl)) ||
            connection.HasIntervalMinutes && connection.IntervalMinutes > MaxIntervalMinutes)
            throw Invalid("Workbook connection metadata exceeds the bounded source-editable profile.", location);
    }

    private static void ValidateBinding(
        SpreadsheetConnectionSourceBinding? desired,
        SpreadsheetConnectionSourceBinding source,
        uint connectionId)
    {
        if (desired is null || !desired.PartPath.Equals(source.PartPath, StringComparison.OrdinalIgnoreCase) ||
            !desired.PartXmlSha256.Equals(source.PartXmlSha256, StringComparison.OrdinalIgnoreCase) ||
            !desired.ConnectionXmlSha256.Equals(source.ConnectionXmlSha256, StringComparison.OrdinalIgnoreCase) ||
            !desired.SemanticSha256.Equals(source.SemanticSha256, StringComparison.OrdinalIgnoreCase) ||
            desired.Editable != source.Editable)
            throw Invalid($"Workbook connection {connectionId} source binding does not match the validated source package.", source.PartPath);
    }

    private static void Patch(XElement element, SpreadsheetConnectionArtifact connection)
    {
        element.SetAttributeValue("name", connection.Name);
        element.SetAttributeValue("description", connection.HasDescription ? connection.Description : null);
        SetOptional(element, "keepAlive", connection.HasKeepAlive, connection.KeepAlive);
        SetOptional(element, "interval", connection.HasIntervalMinutes, connection.IntervalMinutes);
        SetOptional(element, "background", connection.HasBackground, connection.Background);
        SetOptional(element, "refreshOnLoad", connection.HasRefreshOnLoad, connection.RefreshOnLoad);
        SetOptional(element, "saveData", connection.HasSaveData, connection.SaveData);
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

    private static void SetOptional(XElement root, string name, bool hasValue, bool value) =>
        root.SetAttributeValue(name, hasValue ? value ? "1" : "0" : null);

    private static void SetOptional(XElement root, string name, bool hasValue, uint value) =>
        root.SetAttributeValue(name, hasValue ? value.ToString(CultureInfo.InvariantCulture) : null);

    private static string SemanticSha256(SpreadsheetConnectionArtifact connection)
    {
        var values = new[]
        {
            connection.ConnectionId.ToString(CultureInfo.InvariantCulture),
            connection.Name,
            connection.Type.ToString(CultureInfo.InvariantCulture),
            connection.RefreshedVersion.ToString(CultureInfo.InvariantCulture),
            connection.HasDescription ? connection.Description : "<absent>",
            Optional(connection.HasKeepAlive, connection.KeepAlive),
            Optional(connection.HasIntervalMinutes, connection.IntervalMinutes),
            Optional(connection.HasBackground, connection.Background),
            Optional(connection.HasRefreshOnLoad, connection.RefreshOnLoad),
            Optional(connection.HasSaveData, connection.SaveData),
        };
        return Sha256(Encoding.UTF8.GetBytes(string.Join('\0', values)));
    }

    private static string ElementSha256(XElement element) =>
        Sha256(Encoding.UTF8.GetBytes(element.ToString(SaveOptions.DisableFormatting)));

    private static string Optional(bool hasValue, bool value) => hasValue ? value ? "1" : "0" : "<absent>";
    private static string Optional(bool hasValue, uint value) => hasValue ? value.ToString(CultureInfo.InvariantCulture) : "<absent>";
    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static CodecException Invalid(string message, string? location = null) => new("invalid_workbook_connection", message, location);
}
