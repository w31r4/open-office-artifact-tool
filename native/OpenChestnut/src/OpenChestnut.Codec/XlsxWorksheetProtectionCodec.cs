using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns one bounded, passwordless sheetProtection leaf. SpreadsheetML names
// each operation flag in the negative (true means locked); the public wire
// names allowed operations and this codec contains the inversion/defaults.
internal sealed class XlsxWorksheetProtectionCodec
{
    private readonly Worksheet _worksheet;
    private readonly SheetProtection? _element;
    private readonly SpreadsheetWorksheetProtectionArtifact? _sourceArtifact;
    private readonly bool _unsupportedSource;
    private readonly string _partPath;

    internal XlsxWorksheetProtectionCodec(WorksheetPart part)
    {
        _worksheet = part.Worksheet ?? throw Invalid("Worksheet root is missing.", part.Uri.OriginalString);
        _partPath = part.Uri.OriginalString.TrimStart('/');
        var elements = _worksheet.Elements<SheetProtection>().ToArray();
        if (elements.Length == 0) return;
        if (elements.Length != 1 || !TryRead(elements[0], out var artifact))
        {
            _unsupportedSource = true;
            return;
        }
        _element = elements[0];
        artifact!.Source = new SpreadsheetWorksheetProtectionSourceBinding
        {
            PartPath = _partPath,
            WorksheetXmlSha256 = PartSha256(part),
            ProtectionXmlSha256 = ElementSha256(_element),
            SemanticSha256 = SemanticSha256(artifact),
            Editable = true,
        };
        _sourceArtifact = artifact;
    }

    internal SpreadsheetWorksheetProtectionArtifact? Read() => _sourceArtifact?.Clone();

    internal void Apply(SpreadsheetWorksheetProtectionArtifact? desired, bool sourceBound)
    {
        if (desired is not null) Validate(desired);
        if (!sourceBound)
        {
            if (desired?.Enabled == true) _worksheet.AddChild(Create(desired), true);
            return;
        }

        if (_unsupportedSource)
        {
            if (desired is not null)
                throw Unsupported("Source-preserving XLSX export cannot replace worksheet protection that contains password verifiers, extensions, disabled/partial protection, or unsupported markup.");
            return;
        }
        if (_sourceArtifact is null)
        {
            if (desired?.Enabled == true) _worksheet.AddChild(Create(desired), true);
            return;
        }
        if (desired is null) return;
        ValidateBinding(desired.Source, _sourceArtifact.Source!);
        if (!desired.Enabled)
        {
            _element!.Remove();
            return;
        }
        if (SemanticSha256(desired).Equals(_sourceArtifact.Source!.SemanticSha256, StringComparison.OrdinalIgnoreCase)) return;
        _worksheet.ReplaceChild(Create(desired), _element!);
    }

    internal static void Validate(SpreadsheetWorksheetProtectionArtifact? item)
    {
        if (item is null) return;
        if (!item.Enabled && item.AllowedOperations.Count > 0)
            throw Invalid("Disabled worksheet protection cannot declare allowed operations.");
        if (item.AllowedOperations.Count > 15)
            throw Invalid("Worksheet protection contains too many allowed operations.");
        var operations = new HashSet<SpreadsheetWorksheetProtectionOperation>();
        foreach (var operation in item.AllowedOperations)
        {
            if (operation is SpreadsheetWorksheetProtectionOperation.Unspecified ||
                (int)operation < (int)SpreadsheetWorksheetProtectionOperation.SelectLockedCells ||
                (int)operation > (int)SpreadsheetWorksheetProtectionOperation.EditScenarios)
                throw Invalid($"Worksheet protection operation {operation} is unsupported.");
            if (!operations.Add(operation)) throw Invalid($"Worksheet protection operation {operation} is duplicated.");
        }
    }

    private static bool TryRead(SheetProtection element, out SpreadsheetWorksheetProtectionArtifact? artifact)
    {
        artifact = null;
        try
        {
            if (element.HasChildren || element.ExtendedAttributes.Any() ||
                element.Password is not null || element.AlgorithmName is not null || element.HashValue is not null ||
                element.SaltValue is not null || element.SpinCount is not null || element.Sheet?.Value != true) return false;
            var result = new SpreadsheetWorksheetProtectionArtifact { Enabled = true };
            AddAllowed(result, !(element.SelectLockedCells?.Value ?? false), SpreadsheetWorksheetProtectionOperation.SelectLockedCells);
            AddAllowed(result, !(element.SelectUnlockedCells?.Value ?? false), SpreadsheetWorksheetProtectionOperation.SelectUnlockedCells);
            AddAllowed(result, !(element.FormatCells?.Value ?? true), SpreadsheetWorksheetProtectionOperation.FormatCells);
            AddAllowed(result, !(element.FormatColumns?.Value ?? true), SpreadsheetWorksheetProtectionOperation.FormatColumns);
            AddAllowed(result, !(element.FormatRows?.Value ?? true), SpreadsheetWorksheetProtectionOperation.FormatRows);
            AddAllowed(result, !(element.InsertColumns?.Value ?? true), SpreadsheetWorksheetProtectionOperation.InsertColumns);
            AddAllowed(result, !(element.InsertRows?.Value ?? true), SpreadsheetWorksheetProtectionOperation.InsertRows);
            AddAllowed(result, !(element.InsertHyperlinks?.Value ?? true), SpreadsheetWorksheetProtectionOperation.InsertHyperlinks);
            AddAllowed(result, !(element.DeleteColumns?.Value ?? true), SpreadsheetWorksheetProtectionOperation.DeleteColumns);
            AddAllowed(result, !(element.DeleteRows?.Value ?? true), SpreadsheetWorksheetProtectionOperation.DeleteRows);
            AddAllowed(result, !(element.Sort?.Value ?? true), SpreadsheetWorksheetProtectionOperation.Sort);
            AddAllowed(result, !(element.AutoFilter?.Value ?? true), SpreadsheetWorksheetProtectionOperation.AutoFilter);
            AddAllowed(result, !(element.PivotTables?.Value ?? true), SpreadsheetWorksheetProtectionOperation.PivotTables);
            AddAllowed(result, !(element.Objects?.Value ?? false), SpreadsheetWorksheetProtectionOperation.EditObjects);
            AddAllowed(result, !(element.Scenarios?.Value ?? false), SpreadsheetWorksheetProtectionOperation.EditScenarios);
            Validate(result);
            artifact = result;
            return true;
        }
        catch (CodecException)
        {
            return false;
        }
        catch (Exception error) when (error is FormatException or InvalidOperationException or OverflowException)
        {
            return false;
        }
    }

    private static void AddAllowed(SpreadsheetWorksheetProtectionArtifact artifact, bool allowed, SpreadsheetWorksheetProtectionOperation operation)
    {
        if (allowed) artifact.AllowedOperations.Add(operation);
    }

    private static SheetProtection Create(SpreadsheetWorksheetProtectionArtifact item)
    {
        var allowed = item.AllowedOperations.ToHashSet();
        return new SheetProtection
        {
            Sheet = true,
            SelectLockedCells = !allowed.Contains(SpreadsheetWorksheetProtectionOperation.SelectLockedCells),
            SelectUnlockedCells = !allowed.Contains(SpreadsheetWorksheetProtectionOperation.SelectUnlockedCells),
            FormatCells = !allowed.Contains(SpreadsheetWorksheetProtectionOperation.FormatCells),
            FormatColumns = !allowed.Contains(SpreadsheetWorksheetProtectionOperation.FormatColumns),
            FormatRows = !allowed.Contains(SpreadsheetWorksheetProtectionOperation.FormatRows),
            InsertColumns = !allowed.Contains(SpreadsheetWorksheetProtectionOperation.InsertColumns),
            InsertRows = !allowed.Contains(SpreadsheetWorksheetProtectionOperation.InsertRows),
            InsertHyperlinks = !allowed.Contains(SpreadsheetWorksheetProtectionOperation.InsertHyperlinks),
            DeleteColumns = !allowed.Contains(SpreadsheetWorksheetProtectionOperation.DeleteColumns),
            DeleteRows = !allowed.Contains(SpreadsheetWorksheetProtectionOperation.DeleteRows),
            Sort = !allowed.Contains(SpreadsheetWorksheetProtectionOperation.Sort),
            AutoFilter = !allowed.Contains(SpreadsheetWorksheetProtectionOperation.AutoFilter),
            PivotTables = !allowed.Contains(SpreadsheetWorksheetProtectionOperation.PivotTables),
            Objects = !allowed.Contains(SpreadsheetWorksheetProtectionOperation.EditObjects),
            Scenarios = !allowed.Contains(SpreadsheetWorksheetProtectionOperation.EditScenarios),
        };
    }

    private static void ValidateBinding(
        SpreadsheetWorksheetProtectionSourceBinding? desired,
        SpreadsheetWorksheetProtectionSourceBinding source)
    {
        if (desired is null ||
            !desired.PartPath.Equals(source.PartPath, StringComparison.OrdinalIgnoreCase) ||
            !desired.WorksheetXmlSha256.Equals(source.WorksheetXmlSha256, StringComparison.OrdinalIgnoreCase) ||
            !desired.ProtectionXmlSha256.Equals(source.ProtectionXmlSha256, StringComparison.OrdinalIgnoreCase) ||
            !desired.SemanticSha256.Equals(source.SemanticSha256, StringComparison.OrdinalIgnoreCase) ||
            desired.Editable != source.Editable)
            throw new CodecException(
                "invalid_worksheet_protection_source_binding",
                "Worksheet protection source binding does not match the validated source worksheet.",
                source.PartPath);
    }

    private static string SemanticSha256(SpreadsheetWorksheetProtectionArtifact item) => Sha256(Encoding.UTF8.GetBytes(string.Join("\0",
        $"enabled:{item.Enabled}",
        string.Join(",", item.AllowedOperations.OrderBy(operation => (int)operation).Select(operation => ((int)operation).ToString())))));
    private static string ElementSha256(SheetProtection element) => Sha256(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string PartSha256(WorksheetPart part)
    {
        using var source = part.GetStream(FileMode.Open, FileAccess.Read);
        return Convert.ToHexString(SHA256.HashData(source)).ToLowerInvariant();
    }
    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static CodecException Invalid(string message, string? path = null) => new("invalid_worksheet_protection", message, path ?? "artifact.workbook.worksheets.protection");
    private CodecException Unsupported(string message) => new("unsupported_worksheet_protection_edit", message, _partPath);
}
