using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns activeTab plus the workbookViewId=0 worksheet tabSelected group for one
// recognized workbook window. Every other workbook/sheet-view attribute and
// child remains in the source package. Multi-window or incomplete worksheet-
// view graphs stay readable only as opaque source state and fail closed on a
// semantic edit.
internal sealed class XlsxWorkbookViewCodec
{
    private readonly WorkbookPart _workbookPart;
    private readonly Workbook _workbook;
    private readonly Sheet[] _sourceSheets;
    private readonly ViewEntry[] _views;
    private readonly HashSet<int> _protectedSelectedOrdinals = [];
    private readonly ViewEntry? _sourceEntry;
    private readonly SelectionEntry[] _selectionEntries = [];
    private readonly SpreadsheetWorkbookViewArtifact? _sourceArtifact;

    private sealed record ViewEntry(WorkbookView Element, int ActiveOrdinal, bool ActiveWasVisible);
    private sealed record SelectionEntry(
        WorksheetPart Part,
        SheetView Element,
        string WorksheetId,
        uint ViewOrdinal,
        uint WorkbookViewId,
        bool? TabSelected);

    internal XlsxWorkbookViewCodec(WorkbookPart workbookPart, IReadOnlyList<WorksheetArtifact> worksheets)
    {
        _workbookPart = workbookPart;
        _workbook = workbookPart.Workbook ?? throw Invalid("Workbook root is missing.");
        _sourceSheets = _workbook.Sheets?.Elements<Sheet>().ToArray() ?? [];
        _views = (_workbook.BookViews?.Elements<WorkbookView>() ?? [])
            .Select(view => ReadEntry(view, _sourceSheets))
            .ToArray();
        ReadProtectedSelections();
        if (_views.Length != 1) return;

        var entry = _views[0];
        if (entry.ActiveOrdinal < 0 || entry.ActiveOrdinal >= worksheets.Count || !entry.ActiveWasVisible) return;
        var activeWorksheetId = worksheets[entry.ActiveOrdinal].Id;
        var selectionEditable = TryReadSelectionEntries(worksheets, out var selectionEntries, out var selectedWorksheetIds);
        _selectionEntries = selectionEntries;
        var semanticSelectedIds = selectionEditable ? selectedWorksheetIds : [];
        var artifact = new SpreadsheetWorkbookViewArtifact
        {
            ActiveWorksheetId = activeWorksheetId,
            Source = new SpreadsheetWorkbookViewSourceBinding
            {
                Ordinal = 0,
                WorkbookXmlSha256 = PartSha256(workbookPart),
                ViewXmlSha256 = ElementSha256(entry.Element),
                SemanticSha256 = SemanticSha256(activeWorksheetId, semanticSelectedIds),
                Editable = selectionEditable,
            },
        };
        if (selectionEditable)
        {
            artifact.SelectedWorksheetIds.Add(selectedWorksheetIds);
            artifact.Source.WorksheetViews.Add(selectionEntries.Select(SourceBinding));
        }
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
        var selectedWorksheetIds = SelectedWorksheetIds(desired, worksheets, _sourceArtifact.ActiveWorksheetId);
        ValidateBinding(desired.Source, _sourceArtifact.Source);
        var desiredSemantic = SemanticSha256(desired.ActiveWorksheetId, selectedWorksheetIds);
        if (!_sourceArtifact.Source.Editable)
        {
            if (!desiredSemantic.Equals(_sourceArtifact.Source.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                throw Invalid("Source-preserving XLSX export cannot edit an incomplete or opaque worksheet-selection profile.");
            ValidateOpaqueVisibilityTransitions(worksheets);
            return;
        }
        if (desiredSemantic.Equals(_sourceArtifact.Source.SemanticSha256, StringComparison.OrdinalIgnoreCase)) return;

        _sourceEntry.Element.ActiveTab = checked((uint)activeOrdinal);
        PatchSelectionGroup(selectedWorksheetIds);
    }

    internal static void ValidateArtifact(SpreadsheetWorkbookViewArtifact? view, IReadOnlyList<WorksheetArtifact> worksheets)
    {
        if (view is null) return;
        _ = ActiveOrdinal(view, worksheets);
        var legacyPreviousActive = view.Source is not null && view.SelectedWorksheetIds.Count == 1
            ? view.SelectedWorksheetIds[0]
            : null;
        _ = SelectedWorksheetIds(view, worksheets, legacyPreviousActive);
    }

    private void ConfigureSourceFree(SpreadsheetWorkbookViewArtifact? desired, IReadOnlyList<WorksheetArtifact> worksheets)
    {
        var activeOrdinal = desired is null
            ? worksheets.Select((sheet, index) => (sheet, index)).First(item => IsVisible(item.sheet)).index
            : ActiveOrdinal(desired, worksheets);
        var selectedWorksheetIds = desired is null
            ? new[] { worksheets[activeOrdinal].Id }
            : SelectedWorksheetIds(desired, worksheets);
        var views = new BookViews(new WorkbookView { ActiveTab = checked((uint)activeOrdinal) });
        if (_workbook.Sheets is { } sheets) _workbook.InsertBefore(views, sheets);
        else _workbook.Append(views);

        var selected = selectedWorksheetIds.ToHashSet(StringComparer.Ordinal);
        for (var index = 0; index < _sourceSheets.Length; index++)
        {
            var part = WorksheetPartAt(index) ?? throw Invalid($"Worksheet {worksheets[index].Name} has no readable Worksheet part.");
            var sheetViews = part.Worksheet?.SheetViews?.Elements<SheetView>().ToArray() ?? [];
            var matches = sheetViews.Select((view, ordinal) => (view, ordinal))
                .Where(item => (item.view.WorkbookViewId?.Value ?? uint.MaxValue) == 0U)
                .ToArray();
            if (matches.Length != 1)
                throw Invalid($"Worksheet {worksheets[index].Name} must contain exactly one primary sheetView for source-free selection authoring.");
            matches[0].view.TabSelected = selected.Contains(worksheets[index].Id) ? true : null;
        }
    }

    private void PatchSelectionGroup(IReadOnlyList<string> selectedWorksheetIds)
    {
        var selected = selectedWorksheetIds.ToHashSet(StringComparer.Ordinal);
        foreach (var entry in _selectionEntries)
        {
            var shouldSelect = selected.Contains(entry.WorksheetId);
            if (shouldSelect && entry.TabSelected != true) entry.Element.TabSelected = true;
            else if (!shouldSelect && entry.TabSelected == true) entry.Element.RemoveAttribute("tabSelected", string.Empty);
        }
    }

    private bool TryReadSelectionEntries(
        IReadOnlyList<WorksheetArtifact> worksheets,
        out SelectionEntry[] entries,
        out string[] selectedWorksheetIds)
    {
        entries = [];
        selectedWorksheetIds = [];
        if (_sourceSheets.Length != worksheets.Count) return false;
        var collected = new List<SelectionEntry>();
        var selected = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < worksheets.Count; index++)
        {
            var part = WorksheetPartAt(index);
            var views = part?.Worksheet?.SheetViews?.Elements<SheetView>().ToArray() ?? [];
            if (part is null || views.Length != 1 || (views[0].WorkbookViewId?.Value ?? uint.MaxValue) != 0U ||
                !TryReadTabSelected(views[0], out var tabSelected)) return false;
            var entry = new SelectionEntry(part, views[0], worksheets[index].Id, 0, 0, tabSelected);
            collected.Add(entry);
            if (tabSelected == true) selected.Add(worksheets[index].Id);
        }

        var activeWorksheetId = worksheets[_views[0].ActiveOrdinal].Id;
        selected.Add(activeWorksheetId);
        var ordered = worksheets.Where(sheet => selected.Contains(sheet.Id)).Select(sheet => sheet.Id).ToArray();
        if (worksheets.Select((sheet, index) => (sheet, index))
            .Where(item => selected.Contains(item.sheet.Id))
            .Any(item => !IsSourceVisible(_sourceSheets[item.index]))) return false;
        entries = collected.ToArray();
        selectedWorksheetIds = ordered;
        return true;
    }

    private SpreadsheetWorksheetViewSourceBinding SourceBinding(SelectionEntry entry)
    {
        var binding = new SpreadsheetWorksheetViewSourceBinding
        {
            WorksheetId = entry.WorksheetId,
            PartPath = entry.Part.Uri.ToString().TrimStart('/'),
            ViewOrdinal = entry.ViewOrdinal,
            WorkbookViewId = entry.WorkbookViewId,
            WorksheetXmlSha256 = PartSha256(entry.Part),
            ViewXmlSha256 = ElementSha256(entry.Element),
            Editable = true,
        };
        if (entry.TabSelected.HasValue) binding.TabSelected = entry.TabSelected.Value;
        return binding;
    }

    private void ReadProtectedSelections()
    {
        foreach (var view in _views)
            if (view.ActiveWasVisible && view.ActiveOrdinal >= 0 && view.ActiveOrdinal < _sourceSheets.Length)
                _protectedSelectedOrdinals.Add(view.ActiveOrdinal);
        for (var index = 0; index < _sourceSheets.Length; index++)
        {
            var views = WorksheetPartAt(index)?.Worksheet?.SheetViews?.Elements<SheetView>() ?? [];
            foreach (var view in views)
                if (TryReadTabSelected(view, out var selected) && selected == true &&
                    (view.WorkbookViewId?.Value ?? uint.MaxValue) < _views.Length && IsSourceVisible(_sourceSheets[index]))
                    _protectedSelectedOrdinals.Add(index);
        }
    }

    private void ValidateOpaqueVisibilityTransitions(IReadOnlyList<WorksheetArtifact> worksheets)
    {
        foreach (var ordinal in _protectedSelectedOrdinals)
            if (ordinal < worksheets.Count && !IsVisible(worksheets[ordinal]))
                throw Invalid("Source-preserving XLSX export cannot hide an active worksheet or selected worksheet in an opaque or multi-window workbook-view profile.");
    }

    private WorksheetPart? WorksheetPartAt(int ordinal)
    {
        if (ordinal < 0 || ordinal >= _sourceSheets.Length || _sourceSheets[ordinal].Id?.Value is not { Length: > 0 } relationshipId) return null;
        try { return _workbookPart.GetPartById(relationshipId) as WorksheetPart; }
        catch (ArgumentOutOfRangeException) { return null; }
    }

    private static ViewEntry ReadEntry(WorkbookView view, IReadOnlyList<Sheet> sheets)
    {
        var activeOrdinal = checked((int)(view.ActiveTab?.Value ?? 0U));
        var activeWasVisible = activeOrdinal >= 0 && activeOrdinal < sheets.Count && IsSourceVisible(sheets[activeOrdinal]);
        return new ViewEntry(view, activeOrdinal, activeWasVisible);
    }

    private static bool TryReadTabSelected(SheetView view, out bool? selected)
    {
        var attribute = view.GetAttributes().FirstOrDefault(item => item.LocalName.Equals("tabSelected", StringComparison.Ordinal));
        if (string.IsNullOrEmpty(attribute.LocalName))
        {
            selected = null;
            return true;
        }
        selected = attribute.Value switch
        {
            "1" or "true" => true,
            "0" or "false" => false,
            _ => null,
        };
        return selected.HasValue;
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

    private static string[] SelectedWorksheetIds(
        SpreadsheetWorkbookViewArtifact view,
        IReadOnlyList<WorksheetArtifact> worksheets,
        string? legacyPreviousActiveWorksheetId = null)
    {
        var requested = view.SelectedWorksheetIds.Count == 0 ? [view.ActiveWorksheetId] : view.SelectedWorksheetIds.ToArray();
        // The first active-view wire slice exposed only active_worksheet_id.
        // Preserve that editing contract when an imported single selection is
        // unchanged on the wire but its active worksheet has moved.
        if (!string.IsNullOrEmpty(legacyPreviousActiveWorksheetId) &&
            !view.ActiveWorksheetId.Equals(legacyPreviousActiveWorksheetId, StringComparison.Ordinal) &&
            requested.Length == 1 && requested[0].Equals(legacyPreviousActiveWorksheetId, StringComparison.Ordinal))
            requested = [view.ActiveWorksheetId];
        if (requested.Length == 0) throw Invalid("Workbook view must select at least one worksheet.");
        var unique = new HashSet<string>(StringComparer.Ordinal);
        foreach (var id in requested)
        {
            if (!unique.Add(id)) throw Invalid($"Workbook view selected worksheet {id} is duplicated.");
            var matches = worksheets.Where(sheet => sheet.Id.Equals(id, StringComparison.Ordinal)).ToArray();
            if (matches.Length != 1) throw Invalid($"Workbook view selected worksheet {id} does not resolve to exactly one worksheet.");
            if (!IsVisible(matches[0])) throw Invalid($"Workbook view selected worksheet {matches[0].Name} must be visible.");
        }
        if (!unique.Contains(view.ActiveWorksheetId))
            throw Invalid("Workbook view selected worksheets must include the active worksheet.");
        return worksheets.Where(sheet => unique.Contains(sheet.Id)).Select(sheet => sheet.Id).ToArray();
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
            desired.Editable != source.Editable || desired.WorksheetViews.Count != source.WorksheetViews.Count)
            throw Invalid("Workbook-view source binding does not match the validated source workbook.");
        for (var index = 0; index < source.WorksheetViews.Count; index++)
        {
            var target = desired.WorksheetViews[index];
            var original = source.WorksheetViews[index];
            if (!target.WorksheetId.Equals(original.WorksheetId, StringComparison.Ordinal) ||
                !target.PartPath.Equals(original.PartPath, StringComparison.Ordinal) ||
                target.ViewOrdinal != original.ViewOrdinal || target.WorkbookViewId != original.WorkbookViewId ||
                !target.WorksheetXmlSha256.Equals(original.WorksheetXmlSha256, StringComparison.OrdinalIgnoreCase) ||
                !target.ViewXmlSha256.Equals(original.ViewXmlSha256, StringComparison.OrdinalIgnoreCase) ||
                target.HasTabSelected != original.HasTabSelected || target.HasTabSelected && target.TabSelected != original.TabSelected ||
                target.Editable != original.Editable)
                throw Invalid("Workbook-view worksheet source binding does not match the validated source workbook.");
        }
    }

    private static string SemanticSha256(string activeWorksheetId, IReadOnlyList<string> selectedWorksheetIds)
    {
        var selected = selectedWorksheetIds.Count == 0 ? [activeWorksheetId] : selectedWorksheetIds;
        var semantic = new[] { $"activeWorksheetId:{activeWorksheetId}" }
            .Concat(selected.Select(id => $"selectedWorksheetId:{id}"));
        return Sha256(Encoding.UTF8.GetBytes(string.Join("\0", semantic)));
    }
    private static string ElementSha256(OpenXmlElement element) => Sha256(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string PartSha256(OpenXmlPart part)
    {
        using var source = part.GetStream(FileMode.Open, FileAccess.Read);
        return Convert.ToHexString(SHA256.HashData(source)).ToLowerInvariant();
    }
    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static CodecException Invalid(string message) => new("invalid_workbook_view", message, "xl/workbook.xml");
}
