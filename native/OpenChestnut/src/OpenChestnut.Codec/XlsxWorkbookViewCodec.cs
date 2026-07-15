using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns only activeTab for one recognized workbookView. Every other view
// attribute/child and the complete multi-window graph remain in the source
// package. The wire uses worksheet identity rather than leaking an ordinal as
// durable document identity.
internal sealed class XlsxWorkbookViewCodec
{
    private readonly Workbook _workbook;
    private readonly ViewEntry[] _views;
    private readonly ViewEntry? _sourceEntry;
    private readonly SpreadsheetWorkbookViewArtifact? _sourceArtifact;

    private sealed record ViewEntry(WorkbookView Element, int ActiveOrdinal, bool ActiveWasVisible);

    internal XlsxWorkbookViewCodec(WorkbookPart workbookPart, IReadOnlyList<WorksheetArtifact> worksheets)
    {
        _workbook = workbookPart.Workbook ?? throw Invalid("Workbook root is missing.");
        var sourceSheets = _workbook.Sheets?.Elements<Sheet>().ToArray() ?? [];
        _views = (_workbook.BookViews?.Elements<WorkbookView>() ?? [])
            .Select(view => ReadEntry(view, sourceSheets))
            .ToArray();
        if (_views.Length != 1) return;

        var entry = _views[0];
        if (entry.ActiveOrdinal < 0 || entry.ActiveOrdinal >= worksheets.Count || !entry.ActiveWasVisible) return;
        var artifact = new SpreadsheetWorkbookViewArtifact
        {
            ActiveWorksheetId = worksheets[entry.ActiveOrdinal].Id,
            Source = new SpreadsheetWorkbookViewSourceBinding
            {
                Ordinal = 0,
                WorkbookXmlSha256 = PartSha256(workbookPart),
                ViewXmlSha256 = ElementSha256(entry.Element),
                SemanticSha256 = SemanticSha256(worksheets[entry.ActiveOrdinal].Id),
                Editable = true,
            },
        };
        _sourceEntry = entry;
        _sourceArtifact = artifact;
    }

    internal SpreadsheetWorkbookViewArtifact? Read() => _sourceArtifact?.Clone();

    internal void Apply(SpreadsheetWorkbookViewArtifact? desired, bool sourceBound, IReadOnlyList<WorksheetArtifact> worksheets)
    {
        if (!sourceBound)
        {
            ConfigureSourceFree(desired, worksheets);
            return;
        }

        if (_sourceArtifact is null || _sourceEntry is null)
        {
            if (desired is not null)
                throw Invalid(_views.Length == 0
                    ? "Source-preserving XLSX export cannot add a workbook view that was absent from the imported workbook."
                    : "Source-preserving XLSX export cannot replace an opaque or multi-window workbook-view profile.");
            ValidateOpaqueVisibilityTransitions(worksheets);
            return;
        }
        if (desired is null)
            throw Invalid("Source-preserving XLSX export cannot remove the imported workbook view.");

        var activeOrdinal = ActiveOrdinal(desired, worksheets);
        ValidateBinding(desired.Source, _sourceArtifact.Source);
        if (!SemanticSha256(desired.ActiveWorksheetId).Equals(_sourceArtifact.Source.SemanticSha256, StringComparison.OrdinalIgnoreCase))
            _sourceEntry.Element.ActiveTab = checked((uint)activeOrdinal);
    }

    internal static void ValidateArtifact(SpreadsheetWorkbookViewArtifact? view, IReadOnlyList<WorksheetArtifact> worksheets)
    {
        if (view is not null) _ = ActiveOrdinal(view, worksheets);
    }

    private void ConfigureSourceFree(SpreadsheetWorkbookViewArtifact? desired, IReadOnlyList<WorksheetArtifact> worksheets)
    {
        var activeOrdinal = desired is null
            ? worksheets.Select((sheet, index) => (sheet, index)).First(item => IsVisible(item.sheet)).index
            : ActiveOrdinal(desired, worksheets);
        var views = new BookViews(new WorkbookView { ActiveTab = checked((uint)activeOrdinal) });
        if (_workbook.Sheets is { } sheets) _workbook.InsertBefore(views, sheets);
        else _workbook.Append(views);
    }

    private void ValidateOpaqueVisibilityTransitions(IReadOnlyList<WorksheetArtifact> worksheets)
    {
        foreach (var view in _views)
        {
            if (!view.ActiveWasVisible || view.ActiveOrdinal < 0 || view.ActiveOrdinal >= worksheets.Count) continue;
            if (!IsVisible(worksheets[view.ActiveOrdinal]))
                throw Invalid("Source-preserving XLSX export cannot hide an active worksheet in an opaque or multi-window workbook-view profile.");
        }
    }

    private static ViewEntry ReadEntry(WorkbookView view, IReadOnlyList<Sheet> sheets)
    {
        var activeOrdinal = checked((int)(view.ActiveTab?.Value ?? 0U));
        var activeWasVisible = activeOrdinal >= 0 && activeOrdinal < sheets.Count && IsSourceVisible(sheets[activeOrdinal]);
        return new ViewEntry(view, activeOrdinal, activeWasVisible);
    }

    private static bool IsSourceVisible(Sheet sheet)
    {
        var rawState = sheet.GetAttributes().FirstOrDefault(attribute => attribute.LocalName.Equals("state", StringComparison.Ordinal)).Value ?? string.Empty;
        return rawState is "" or "visible";
    }

    private static int ActiveOrdinal(SpreadsheetWorkbookViewArtifact view, IReadOnlyList<WorksheetArtifact> worksheets)
    {
        if (string.IsNullOrWhiteSpace(view.ActiveWorksheetId))
            throw Invalid("Workbook view must identify an active worksheet.");
        var matches = worksheets.Select((sheet, index) => (sheet, index))
            .Where(item => item.sheet.Id.Equals(view.ActiveWorksheetId, StringComparison.Ordinal))
            .ToArray();
        if (matches.Length != 1)
            throw Invalid($"Workbook view active worksheet {view.ActiveWorksheetId} does not resolve to exactly one worksheet.");
        if (!IsVisible(matches[0].sheet))
            throw Invalid($"Workbook view active worksheet {matches[0].sheet.Name} must be visible.");
        return matches[0].index;
    }

    private static bool IsVisible(WorksheetArtifact sheet) =>
        sheet.Visibility == SpreadsheetWorksheetVisibility.Visible ||
        sheet.Visibility == SpreadsheetWorksheetVisibility.Unspecified && sheet.Source is null;

    private static void ValidateBinding(SpreadsheetWorkbookViewSourceBinding? desired, SpreadsheetWorkbookViewSourceBinding source)
    {
        if (desired is null || desired.Ordinal != 0 || desired.Ordinal != source.Ordinal ||
            !desired.WorkbookXmlSha256.Equals(source.WorkbookXmlSha256, StringComparison.OrdinalIgnoreCase) ||
            !desired.ViewXmlSha256.Equals(source.ViewXmlSha256, StringComparison.OrdinalIgnoreCase) ||
            !desired.SemanticSha256.Equals(source.SemanticSha256, StringComparison.OrdinalIgnoreCase) ||
            desired.Editable != source.Editable)
            throw Invalid("Workbook-view source binding does not match the validated source workbook.");
    }

    private static string SemanticSha256(string activeWorksheetId) =>
        Sha256(Encoding.UTF8.GetBytes($"activeWorksheetId:{activeWorksheetId}"));
    private static string ElementSha256(WorkbookView element) => Sha256(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string PartSha256(WorkbookPart part)
    {
        using var source = part.GetStream(FileMode.Open, FileAccess.Read);
        return Convert.ToHexString(SHA256.HashData(source)).ToLowerInvariant();
    }
    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static CodecException Invalid(string message) => new("invalid_workbook_view", message, "xl/workbook.xml");
}
