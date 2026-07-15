using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns the bounded workbook.xml <sheet> name/visibility projection. Package
// locators stay immutable and unknown state values remain opaque. Workbook
// view selection is independently owned by XlsxWorkbookViewCodec.
internal sealed class XlsxWorksheetMetadataCodec
{
    private readonly Entry[] _entries;

    private sealed record Entry(
        Sheet Element,
        string Name,
        SpreadsheetWorksheetVisibility? Visibility,
        SpreadsheetWorksheetSourceBinding Binding);

    internal XlsxWorksheetMetadataCodec(WorkbookPart workbookPart)
    {
        var workbook = workbookPart.Workbook ?? throw Invalid("Workbook root is missing.");
        var sheets = workbook.Sheets?.Elements<Sheet>().ToArray() ?? [];
        _entries = sheets.Select((element, index) => ReadEntry(element, checked((uint)index))).ToArray();
    }

    internal void ReadInto(WorksheetArtifact target, int index)
    {
        var entry = _entries[index];
        target.Visibility = entry.Visibility ?? SpreadsheetWorksheetVisibility.Unspecified;
        target.Source = entry.Binding.Clone();
    }

    internal void Apply(IReadOnlyList<WorksheetArtifact> desired)
    {
        if (desired.Count != _entries.Length)
            throw Invalid("Source-preserving XLSX export requires the imported worksheet count and order to remain unchanged.");

        for (var index = 0; index < desired.Count; index++)
        {
            var source = _entries[index];
            var target = desired[index];
            ValidateBinding(target.Source, source.Binding, checked((uint)index));
            if (source.Visibility is null)
            {
                if (!target.Name.Equals(source.Name, StringComparison.Ordinal) || target.Visibility != SpreadsheetWorksheetVisibility.Unspecified)
                    throw Invalid($"Source-preserving XLSX export cannot replace opaque worksheet metadata at ordinal {index}.");
                continue;
            }

            ValidateVisibility(target.Visibility, target.Name);
            if (SemanticSha256(target.Name, target.Visibility).Equals(source.Binding.SemanticSha256, StringComparison.OrdinalIgnoreCase)) continue;
            source.Element.Name = target.Name;
            if (target.Visibility != source.Visibility)
                source.Element.State = target.Visibility switch
                {
                    SpreadsheetWorksheetVisibility.Visible => SheetStateValues.Visible,
                    SpreadsheetWorksheetVisibility.Hidden => SheetStateValues.Hidden,
                    SpreadsheetWorksheetVisibility.VeryHidden => SheetStateValues.VeryHidden,
                    _ => throw Invalid($"Worksheet {target.Name} has an unsupported visibility value."),
                };
        }

        ValidateVisibleSheet(desired);
    }

    internal static Sheet Create(WorksheetArtifact source, uint ordinal, string relationshipId)
    {
        var visibility = EffectiveVisibility(source);
        ValidateVisibility(visibility, source.Name);
        var sheet = new Sheet
        {
            Id = relationshipId,
            SheetId = ordinal + 1,
            Name = source.Name,
        };
        if (visibility != SpreadsheetWorksheetVisibility.Visible)
            sheet.State = visibility == SpreadsheetWorksheetVisibility.Hidden ? SheetStateValues.Hidden : SheetStateValues.VeryHidden;
        return sheet;
    }

    internal static void ValidateArtifact(IReadOnlyList<WorksheetArtifact> worksheets)
    {
        foreach (var sheet in worksheets)
        {
            if (sheet.Visibility == SpreadsheetWorksheetVisibility.Unspecified && sheet.Source is { Editable: false }) continue;
            ValidateVisibility(EffectiveVisibility(sheet), sheet.Name);
        }
        ValidateVisibleSheet(worksheets);
    }

    private static Entry ReadEntry(Sheet element, uint ordinal)
    {
        var name = element.Name?.Value ?? string.Empty;
        var relationshipId = element.Id?.Value ?? string.Empty;
        var sheetId = element.SheetId?.Value ?? 0U;
        var rawState = element.GetAttributes().FirstOrDefault(attribute => attribute.LocalName.Equals("state", StringComparison.Ordinal)).Value ?? string.Empty;
        SpreadsheetWorksheetVisibility? visibility = rawState switch
        {
            "" or "visible" => SpreadsheetWorksheetVisibility.Visible,
            "hidden" => SpreadsheetWorksheetVisibility.Hidden,
            "veryHidden" => SpreadsheetWorksheetVisibility.VeryHidden,
            _ => null,
        };
        var binding = new SpreadsheetWorksheetSourceBinding
        {
            Ordinal = ordinal,
            RelationshipId = relationshipId,
            SheetId = sheetId,
            SheetElementSha256 = Sha256(Encoding.UTF8.GetBytes(element.OuterXml)),
            SemanticSha256 = visibility is null ? SemanticSha256(name, rawState) : SemanticSha256(name, visibility.Value),
            Editable = visibility is not null,
        };
        return new Entry(element, name, visibility, binding);
    }

    private static void ValidateVisibility(SpreadsheetWorksheetVisibility visibility, string name)
    {
        if (visibility is not (SpreadsheetWorksheetVisibility.Visible or SpreadsheetWorksheetVisibility.Hidden or SpreadsheetWorksheetVisibility.VeryHidden))
            throw Invalid($"Worksheet {name} visibility must be visible, hidden, or veryHidden.");
    }

    private static void ValidateVisibleSheet(IReadOnlyList<WorksheetArtifact> worksheets)
    {
        if (!worksheets.Any(sheet => EffectiveVisibility(sheet) == SpreadsheetWorksheetVisibility.Visible))
            throw Invalid("Workbook must contain at least one visible worksheet.");
    }

    private static SpreadsheetWorksheetVisibility EffectiveVisibility(WorksheetArtifact sheet) =>
        sheet.Visibility == SpreadsheetWorksheetVisibility.Unspecified && sheet.Source is null
            ? SpreadsheetWorksheetVisibility.Visible
            : sheet.Visibility;

    private static void ValidateBinding(SpreadsheetWorksheetSourceBinding? desired, SpreadsheetWorksheetSourceBinding source, uint ordinal)
    {
        if (desired is null || desired.Ordinal != ordinal || desired.Ordinal != source.Ordinal ||
            !desired.RelationshipId.Equals(source.RelationshipId, StringComparison.Ordinal) || desired.SheetId != source.SheetId ||
            !desired.SheetElementSha256.Equals(source.SheetElementSha256, StringComparison.OrdinalIgnoreCase) ||
            !desired.SemanticSha256.Equals(source.SemanticSha256, StringComparison.OrdinalIgnoreCase) || desired.Editable != source.Editable)
            throw Invalid($"Worksheet source binding at ordinal {ordinal} does not match the validated source workbook.");
    }

    private static string SemanticSha256(string name, SpreadsheetWorksheetVisibility visibility) => SemanticSha256(name, visibility.ToString());
    private static string SemanticSha256(string name, string state) => Sha256(Encoding.UTF8.GetBytes(string.Join("\0", $"name:{name}", $"visibility:{state}")));
    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static CodecException Invalid(string message) => new("invalid_worksheet_metadata", message, "xl/workbook.xml");
}
