using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns each workbookView activeTab plus the matching worksheet sheetView
// tabSelected group. The public wire keeps the primary window in the original
// field and appends later windows separately. Geometry, zoom, panes, first-sheet
// scrolling, visibility, and extensions remain source-owned.
internal sealed class XlsxWorkbookViewCodec
{
    private readonly WorkbookPart _workbookPart;
    private readonly Workbook _workbook;
    private readonly Sheet[] _sourceSheets;
    private readonly ViewEntry[] _views;
    private readonly HashSet<int> _protectedSelectedOrdinals = [];
    private readonly SelectionEntry[][] _selectionEntriesByView = [];
    private readonly SpreadsheetWorkbookViewArtifact[] _sourceArtifacts = [];

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
        if (_views.Length == 0 || _views.Any(entry => entry.ActiveOrdinal < 0 || entry.ActiveOrdinal >= worksheets.Count || !entry.ActiveWasVisible)) return;

        var complete = TryReadCompleteSelectionGraph(worksheets, out var selectionEntries, out var selectedWorksheetIds);
        if (_views.Length > 1 && !complete) return;

        _selectionEntriesByView = complete ? selectionEntries : [[]];
        _sourceArtifacts = _views.Select((entry, ordinal) =>
        {
            var activeWorksheetId = worksheets[entry.ActiveOrdinal].Id;
            var selected = complete ? selectedWorksheetIds[ordinal] : [];
            var source = new SpreadsheetWorkbookViewSourceBinding
            {
                Ordinal = checked((uint)ordinal),
                WorkbookXmlSha256 = PartSha256(workbookPart),
                ViewXmlSha256 = ElementSha256(entry.Element),
                SemanticSha256 = SemanticSha256(activeWorksheetId, selected),
                Editable = complete,
            };
            if (complete) source.WorksheetViews.Add(selectionEntries[ordinal].Select(SourceBinding));
            var artifact = new SpreadsheetWorkbookViewArtifact
            {
                ActiveWorksheetId = activeWorksheetId,
                Source = source,
            };
            if (complete) artifact.SelectedWorksheetIds.Add(selected);
            return artifact;
        }).ToArray();
    }

    internal SpreadsheetWorkbookViewArtifact[] Read() => _sourceArtifacts.Select(view => view.Clone()).ToArray();

    internal void Apply(
        SpreadsheetWorkbookViewArtifact? primary,
        IEnumerable<SpreadsheetWorkbookViewArtifact> additional,
        bool sourceBound,
        IReadOnlyList<WorksheetArtifact> worksheets)
    {
        var desired = DesiredViews(primary, additional);
        if (!sourceBound)
        {
            ConfigureSourceFree(desired, worksheets);
            return;
        }

        if (_sourceArtifacts.Length == 0)
        {
            if (desired.Length > 0)
                throw Invalid(_views.Length == 0
                    ? "Source-preserving XLSX export cannot add a workbook view that was absent from the imported workbook."
                    : _views.Length > 1
                        ? "Source-preserving XLSX export cannot replace an opaque or incomplete multi-window workbook-view profile."
                        : "Source-preserving XLSX export cannot replace an opaque workbook-view profile.");
            ValidateOpaqueVisibilityTransitions(worksheets);
            return;
        }
        if (desired.Length == 0) throw Invalid("Source-preserving XLSX export cannot remove the imported workbook views.");
        if (desired.Length != _sourceArtifacts.Length)
            throw Invalid("Source-preserving XLSX export cannot change the imported workbook-window count or order.");

        for (var ordinal = 0; ordinal < desired.Length; ordinal++)
        {
            var target = desired[ordinal];
            var source = _sourceArtifacts[ordinal];
            var activeOrdinal = ActiveOrdinal(target, worksheets);
            var selectedWorksheetIds = SelectedWorksheetIds(target, worksheets, source.ActiveWorksheetId);
            ValidateBinding(target.Source, source.Source);
            var desiredSemantic = SemanticSha256(target.ActiveWorksheetId, selectedWorksheetIds);
            if (!source.Source.Editable)
            {
                if (!desiredSemantic.Equals(source.Source.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                    throw Invalid("Source-preserving XLSX export cannot edit an incomplete or opaque worksheet-selection profile.");
                continue;
            }
            if (desiredSemantic.Equals(source.Source.SemanticSha256, StringComparison.OrdinalIgnoreCase)) continue;

            _views[ordinal].Element.ActiveTab = checked((uint)activeOrdinal);
            PatchSelectionGroup(_selectionEntriesByView[ordinal], selectedWorksheetIds);
        }
    }

    internal static void ValidateArtifact(
        SpreadsheetWorkbookViewArtifact? primary,
        IEnumerable<SpreadsheetWorkbookViewArtifact> additional,
        IReadOnlyList<WorksheetArtifact> worksheets)
    {
        var views = DesiredViews(primary, additional);
        for (var ordinal = 0; ordinal < views.Length; ordinal++)
        {
            _ = ActiveOrdinal(views[ordinal], worksheets);
            var previousActive = views[ordinal].Source is not null && views[ordinal].SelectedWorksheetIds.Count == 1
                ? views[ordinal].SelectedWorksheetIds[0]
                : null;
            _ = SelectedWorksheetIds(views[ordinal], worksheets, previousActive);
            if (views[ordinal].Source is not null && views[ordinal].Source.Ordinal != checked((uint)ordinal))
                throw Invalid($"Workbook view source ordinal {views[ordinal].Source.Ordinal} does not match window position {ordinal}.");
        }
    }

    private static SpreadsheetWorkbookViewArtifact[] DesiredViews(
        SpreadsheetWorkbookViewArtifact? primary,
        IEnumerable<SpreadsheetWorkbookViewArtifact> additional)
    {
        var tail = additional.ToArray();
        if (primary is null)
        {
            if (tail.Length > 0) throw Invalid("Additional workbook windows require a primary workbook view.");
            return [];
        }
        return [primary, .. tail];
    }

    private void ConfigureSourceFree(SpreadsheetWorkbookViewArtifact[] requested, IReadOnlyList<WorksheetArtifact> worksheets)
    {
        SpreadsheetWorkbookViewArtifact[] desired;
        if (requested.Length == 0)
        {
            var active = worksheets.First(IsVisible);
            desired = [new SpreadsheetWorkbookViewArtifact { ActiveWorksheetId = active.Id }];
        }
        else desired = requested;

        var selectedByView = desired.Select(view => SelectedWorksheetIds(view, worksheets)).ToArray();
        var bookViews = new BookViews();
        foreach (var view in desired)
            bookViews.Append(new WorkbookView { ActiveTab = checked((uint)ActiveOrdinal(view, worksheets)) });
        if (_workbook.Sheets is { } sheets) _workbook.InsertBefore(bookViews, sheets);
        else _workbook.Append(bookViews);

        for (var sheetIndex = 0; sheetIndex < _sourceSheets.Length; sheetIndex++)
        {
            var part = WorksheetPartAt(sheetIndex) ?? throw Invalid($"Worksheet {worksheets[sheetIndex].Name} has no readable Worksheet part.");
            var sheetViews = part.Worksheet?.SheetViews ?? throw Invalid($"Worksheet {worksheets[sheetIndex].Name} has no sheetViews collection.");
            var existing = sheetViews.Elements<SheetView>().ToArray();
            if (existing.Length != 1 || (existing[0].WorkbookViewId?.Value ?? uint.MaxValue) != 0U)
                throw Invalid($"Worksheet {worksheets[sheetIndex].Name} must contain exactly one primary sheetView for source-free window authoring.");
            var template = (SheetView)existing[0].CloneNode(true);
            sheetViews.RemoveAllChildren<SheetView>();
            for (var ordinal = 0; ordinal < desired.Length; ordinal++)
            {
                var view = (SheetView)template.CloneNode(true);
                view.WorkbookViewId = checked((uint)ordinal);
                SetTabSelected(view, selectedByView[ordinal].Contains(worksheets[sheetIndex].Id, StringComparer.Ordinal));
                sheetViews.Append(view);
            }
        }
    }

    private static void PatchSelectionGroup(IEnumerable<SelectionEntry> entries, IReadOnlyList<string> selectedWorksheetIds)
    {
        var selected = selectedWorksheetIds.ToHashSet(StringComparer.Ordinal);
        foreach (var entry in entries)
            SetTabSelected(entry.Element, selected.Contains(entry.WorksheetId));
    }

    private static void SetTabSelected(SheetView view, bool selected)
    {
        if (selected) view.TabSelected = true;
        else
        {
            view.TabSelected = null;
            view.RemoveAttribute("tabSelected", string.Empty);
        }
    }

    private bool TryReadCompleteSelectionGraph(
        IReadOnlyList<WorksheetArtifact> worksheets,
        out SelectionEntry[][] entriesByView,
        out string[][] selectedWorksheetIdsByView)
    {
        entriesByView = [];
        selectedWorksheetIdsByView = [];
        if (_sourceSheets.Length != worksheets.Count || _views.Length == 0) return false;
        var collected = Enumerable.Range(0, _views.Length).Select(_ => new List<SelectionEntry>()).ToArray();
        var selected = Enumerable.Range(0, _views.Length).Select(_ => new HashSet<string>(StringComparer.Ordinal)).ToArray();

        for (var sheetIndex = 0; sheetIndex < worksheets.Count; sheetIndex++)
        {
            var part = WorksheetPartAt(sheetIndex);
            var views = part?.Worksheet?.SheetViews?.Elements<SheetView>().ToArray() ?? [];
            if (part is null || views.Length != _views.Length) return false;
            var seen = new HashSet<uint>();
            for (var viewOrdinal = 0; viewOrdinal < views.Length; viewOrdinal++)
            {
                var workbookViewId = views[viewOrdinal].WorkbookViewId?.Value ?? uint.MaxValue;
                if (workbookViewId >= _views.Length || !seen.Add(workbookViewId) || !TryReadTabSelected(views[viewOrdinal], out var tabSelected)) return false;
                var entry = new SelectionEntry(part, views[viewOrdinal], worksheets[sheetIndex].Id, checked((uint)viewOrdinal), workbookViewId, tabSelected);
                collected[workbookViewId].Add(entry);
                if (tabSelected == true) selected[workbookViewId].Add(worksheets[sheetIndex].Id);
            }
        }

        var ordered = new string[_views.Length][];
        for (var ordinal = 0; ordinal < _views.Length; ordinal++)
        {
            if (collected[ordinal].Count != worksheets.Count) return false;
            selected[ordinal].Add(worksheets[_views[ordinal].ActiveOrdinal].Id);
            ordered[ordinal] = worksheets.Where(sheet => selected[ordinal].Contains(sheet.Id)).Select(sheet => sheet.Id).ToArray();
            if (worksheets.Select((sheet, index) => (sheet, index))
                .Where(item => selected[ordinal].Contains(item.sheet.Id))
                .Any(item => !IsSourceVisible(_sourceSheets[item.index]))) return false;
        }
        entriesByView = collected.Select(entries => entries.ToArray()).ToArray();
        selectedWorksheetIdsByView = ordered;
        return true;
    }

    private static SpreadsheetWorksheetViewSourceBinding SourceBinding(SelectionEntry entry)
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
        var rawOrdinal = view.ActiveTab?.Value ?? 0U;
        var activeOrdinal = rawOrdinal <= int.MaxValue ? (int)rawOrdinal : -1;
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
        if (string.IsNullOrWhiteSpace(view.ActiveWorksheetId)) throw Invalid("Workbook view must identify an active worksheet.");
        var matches = worksheets.Select((sheet, index) => (sheet, index))
            .Where(item => item.sheet.Id.Equals(view.ActiveWorksheetId, StringComparison.Ordinal))
            .ToArray();
        if (matches.Length != 1) throw Invalid($"Workbook view active worksheet {view.ActiveWorksheetId} does not resolve to exactly one worksheet.");
        if (!IsVisible(matches[0].sheet)) throw Invalid($"Workbook view active worksheet {matches[0].sheet.Name} must be visible.");
        return matches[0].index;
    }

    private static string[] SelectedWorksheetIds(
        SpreadsheetWorkbookViewArtifact view,
        IReadOnlyList<WorksheetArtifact> worksheets,
        string? legacyPreviousActiveWorksheetId = null)
    {
        var requested = view.SelectedWorksheetIds.Count == 0 ? [view.ActiveWorksheetId] : view.SelectedWorksheetIds.ToArray();
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
        if (!unique.Contains(view.ActiveWorksheetId)) throw Invalid("Workbook view selected worksheets must include the active worksheet.");
        return worksheets.Where(sheet => unique.Contains(sheet.Id)).Select(sheet => sheet.Id).ToArray();
    }

    private static bool IsVisible(WorksheetArtifact sheet) =>
        sheet.Visibility == SpreadsheetWorksheetVisibility.Visible ||
        sheet.Visibility == SpreadsheetWorksheetVisibility.Unspecified && sheet.Source is null;

    private static void ValidateBinding(SpreadsheetWorkbookViewSourceBinding? desired, SpreadsheetWorkbookViewSourceBinding source)
    {
        if (desired is null || desired.Ordinal != source.Ordinal ||
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
