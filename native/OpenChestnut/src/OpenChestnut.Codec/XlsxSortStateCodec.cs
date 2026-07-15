using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Shared SpreadsheetML sortState codec. The same vocabulary appears at the
// worksheet, AutoFilter, and QueryTable refresh boundaries, while callers own
// the context-specific topology and columnSort policy.
internal static class XlsxSortStateCodec
{
    private static readonly XNamespace Spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly HashSet<string> SortMethods = new(StringComparer.Ordinal) { "none", "pinYin", "stroke" };
    private static readonly IReadOnlyDictionary<string, uint> IconSets = new Dictionary<string, uint>(StringComparer.Ordinal)
    {
        ["3Arrows"] = 3, ["3ArrowsGray"] = 3, ["3Flags"] = 3, ["3TrafficLights1"] = 3,
        ["3TrafficLights2"] = 3, ["3Signs"] = 3, ["3Symbols"] = 3, ["3Symbols2"] = 3,
        ["4Arrows"] = 4, ["4ArrowsGray"] = 4, ["4RedToBlack"] = 4, ["4Rating"] = 4,
        ["4TrafficLights"] = 4, ["5Arrows"] = 5, ["5ArrowsGray"] = 5, ["5Rating"] = 5,
        ["5Quarters"] = 5,
    };

    internal static bool TryRead(
        XElement element,
        XlsxCellStyleCodec styles,
        bool allowExtensions,
        out SpreadsheetTableSortStateArtifact? artifact)
    {
        artifact = null;
        if (element.Name != Spreadsheet + "sortState" || element.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration &&
            (attribute.Name.Namespace != XNamespace.None || attribute.Name.LocalName is not ("ref" or "caseSensitive" or "sortMethod" or "columnSort"))) ||
            element.Attribute("ref") is not XAttribute reference ||
            !TryBool(element.Attribute("caseSensitive")?.Value, defaultValue: false, out var caseSensitive) ||
            !TryOptionalBool(element.Attribute("columnSort")?.Value, out var columnSort)) return false;
        var sortMethod = element.Attribute("sortMethod")?.Value;
        if (sortMethod is not null && !SortMethods.Contains(sortMethod)) return false;
        var children = element.Elements().ToArray();
        var extensionIndexes = children.Select((child, index) => (child, index)).Where(item => item.child.Name == Spreadsheet + "extLst").Select(item => item.index).ToArray();
        if (!allowExtensions && extensionIndexes.Length > 0 || extensionIndexes.Length > 1 ||
            extensionIndexes.Length == 1 && extensionIndexes[0] != children.Length - 1 ||
            children.Any(child => child.Name != Spreadsheet + "sortCondition" && child.Name != Spreadsheet + "extLst")) return false;
        var conditions = children.Where(child => child.Name == Spreadsheet + "sortCondition").ToArray();
        if (conditions.Length is < 1 or > 64) return false;
        var sort = new SpreadsheetTableSortStateArtifact { Reference = reference.Value, CaseSensitive = caseSensitive };
        if (sortMethod is not null) sort.SortMethod = sortMethod;
        if (columnSort is not null) sort.ColumnSort = columnSort.Value;
        foreach (var condition in conditions)
        {
            if (condition.Elements().Any() || condition.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration &&
                (attribute.Name.Namespace != XNamespace.None || attribute.Name.LocalName is not ("ref" or "descending" or "sortBy" or "iconSet" or "iconId" or "dxfId" or "customList"))) ||
                condition.Attribute("ref") is not XAttribute conditionReference ||
                !TryBool(condition.Attribute("descending")?.Value, defaultValue: false, out var descending)) return false;
            var sortBy = condition.Attribute("sortBy")?.Value ?? "value";
            var iconSet = condition.Attribute("iconSet")?.Value;
            if (!TryOptionalUInt(condition.Attribute("iconId")?.Value, out var iconId) ||
                !TryOptionalUInt(condition.Attribute("dxfId")?.Value, out var differentialFormatId)) return false;
            var result = new SpreadsheetTableSortConditionArtifact { Reference = conditionReference.Value, Descending = descending };
            var customList = condition.Attribute("customList")?.Value;
            if (sortBy == "icon")
            {
                if (iconSet is null || differentialFormatId is not null || customList is not null || !IconSets.ContainsKey(iconSet)) return false;
                var icon = new SpreadsheetTableIconArtifact { IconSet = iconSet };
                if (iconId is not null) icon.IconId = iconId.Value;
                if (!ValidIcon(icon)) return false;
                result.Icon = icon;
            }
            else if (sortBy is "cellColor" or "fontColor")
            {
                if (iconSet is not null || iconId is not null || differentialFormatId is null || customList is not null ||
                    !styles.TryReadTableColor(differentialFormatId.Value, sortBy == "cellColor", out var color)) return false;
                result.Color = color;
            }
            else if (sortBy != "value" || iconSet is not null || iconId is not null || differentialFormatId is not null) return false;
            if (customList is not null) result.CustomList = customList;
            sort.Conditions.Add(result);
        }
        artifact = sort;
        return true;
    }

    internal static void Validate(
        SpreadsheetTableSortStateArtifact sort,
        (uint Top, uint Left, uint Bottom, uint Right)? containerBounds,
        string location,
        string subject,
        bool allowColumnSort,
        string errorCode = "invalid_worksheet_sort")
    {
        if (!TryRange(sort.Reference, out var bounds))
            throw Invalid(errorCode, $"{subject} sort range must be a valid SpreadsheetML range.", location);
        if (containerBounds is { } container &&
            (bounds.Top < container.Top || bounds.Left < container.Left || bounds.Bottom > container.Bottom || bounds.Right > container.Right))
            throw Invalid(errorCode, $"{subject} sort range must be contained in the source table range.", location);
        if (sort.Conditions.Count is < 1 or > 64)
            throw Invalid(errorCode, $"{subject} sort state must provide between one and 64 conditions.", location);
        if (sort.HasSortMethod && !SortMethods.Contains(sort.SortMethod))
            throw Invalid(errorCode, $"{subject} has an invalid locale-specific sort method.", location);
        if (!allowColumnSort && sort.HasColumnSort)
            throw Invalid(errorCode, $"{subject} cannot define columnSort inside an AutoFilter.", location);

        var columnSort = sort.HasColumnSort && sort.ColumnSort;
        var axes = new HashSet<uint>();
        foreach (var condition in sort.Conditions)
        {
            if (!TryRange(condition.Reference, out var conditionBounds) ||
                (columnSort
                    ? conditionBounds.Top != conditionBounds.Bottom || conditionBounds.Left != bounds.Left || conditionBounds.Right != bounds.Right ||
                      conditionBounds.Top < bounds.Top || conditionBounds.Bottom > bounds.Bottom || !axes.Add(conditionBounds.Top)
                    : conditionBounds.Left != conditionBounds.Right || conditionBounds.Top != bounds.Top || conditionBounds.Bottom != bounds.Bottom ||
                      conditionBounds.Left < bounds.Left || conditionBounds.Right > bounds.Right || !axes.Add(conditionBounds.Left)))
                throw Invalid(errorCode, $"{subject} has an invalid or duplicate {(columnSort ? "row" : "column")} sort condition.", location);
            if (condition.Icon is not null && !ValidIcon(condition.Icon))
                throw Invalid(errorCode, $"{subject} has an invalid icon-sort condition.", location);
            if (condition.Icon is not null && condition.Color is not null)
                throw Invalid(errorCode, $"{subject} sort condition cannot combine icon and color selectors.", location);
            if (condition.HasCustomList && (string.IsNullOrWhiteSpace(condition.CustomList) || condition.CustomList.Length > 32_767 ||
                condition.CustomList.Any(char.IsControl) || condition.Icon is not null || condition.Color is not null))
                throw Invalid(errorCode, $"{subject} has an invalid custom-list value-sort condition.", location);
            if (condition.Color is not null) XlsxCellStyleCodec.ValidateTableColor(condition.Color, subject);
        }
    }

    internal static XElement Create(SpreadsheetTableSortStateArtifact sort, XlsxCellStyleCodec styles)
    {
        var element = new XElement(Spreadsheet + "sortState", new XAttribute("ref", sort.Reference));
        ApplyAttributes(element, sort);
        element.Add(sort.Conditions.Select(condition => CreateCondition(condition, styles)));
        return element;
    }

    internal static void Patch(XElement element, SpreadsheetTableSortStateArtifact sort, XlsxCellStyleCodec styles)
    {
        element.SetAttributeValue("ref", sort.Reference);
        ApplyAttributes(element, sort);
        var conditions = element.Elements(Spreadsheet + "sortCondition").ToArray();
        if (conditions.Length != sort.Conditions.Count)
            throw Invalid("invalid_worksheet_sort", "Source-preserving XLSX export cannot change sort condition topology.", element.Document?.BaseUri);
        for (var index = 0; index < conditions.Length; index++) ApplyCondition(conditions[index], sort.Conditions[index], styles);
    }

    internal static IEnumerable<string> Semantics(SpreadsheetTableSortStateArtifact? sort, string absent)
    {
        if (sort is null) { yield return absent; yield break; }
        yield return sort.Reference;
        yield return sort.CaseSensitive.ToString();
        yield return sort.HasSortMethod ? sort.SortMethod : "no-sort-method";
        yield return sort.HasColumnSort ? sort.ColumnSort.ToString() : "no-column-sort";
        foreach (var condition in sort.Conditions)
        {
            yield return condition.Reference;
            yield return condition.Descending.ToString();
            yield return condition.Icon?.IconSet ?? "value";
            yield return condition.Icon is { HasIconId: true } ? condition.Icon.IconId.ToString(CultureInfo.InvariantCulture) : "no-icon-id";
            yield return condition.Color?.TargetCase.ToString() ?? "no-color";
            yield return condition.Color?.Color?.SourceCase.ToString() ?? "no-color-source";
            yield return condition.Color?.Color?.SourceCase switch
            {
                SpreadsheetColor.SourceOneofCase.Rgb => condition.Color.Color.Rgb,
                SpreadsheetColor.SourceOneofCase.Theme => condition.Color.Color.Theme.ToString(CultureInfo.InvariantCulture),
                SpreadsheetColor.SourceOneofCase.Indexed => condition.Color.Color.Indexed.ToString(CultureInfo.InvariantCulture),
                SpreadsheetColor.SourceOneofCase.Automatic => condition.Color.Color.Automatic.ToString(),
                _ => string.Empty,
            };
            yield return condition.Color?.Color is { HasTint: true } color ? color.Tint.ToString("R", CultureInfo.InvariantCulture) : "no-tint";
            yield return condition.HasCustomList ? condition.CustomList : "no-custom-list";
        }
    }

    internal static bool TryRange(string reference, out (uint Top, uint Left, uint Bottom, uint Right) bounds)
    {
        bounds = default;
        var parts = reference.Split(':');
        if (parts.Length is < 1 or > 2 || !TryCell(parts[0], out var first) ||
            !TryCell(parts.Length == 2 ? parts[1] : parts[0], out var second) ||
            first.Row > second.Row || first.Column > second.Column) return false;
        bounds = (first.Row, first.Column, second.Row, second.Column);
        return true;
    }

    private static XElement CreateCondition(SpreadsheetTableSortConditionArtifact condition, XlsxCellStyleCodec styles)
    {
        var element = new XElement(Spreadsheet + "sortCondition");
        ApplyCondition(element, condition, styles);
        return element;
    }

    private static void ApplyAttributes(XElement element, SpreadsheetTableSortStateArtifact sort)
    {
        element.SetAttributeValue("caseSensitive", sort.CaseSensitive ? "1" : null);
        element.SetAttributeValue("sortMethod", sort.HasSortMethod ? sort.SortMethod : null);
        element.SetAttributeValue("columnSort", sort.HasColumnSort ? sort.ColumnSort ? "1" : "0" : null);
    }

    private static void ApplyCondition(XElement element, SpreadsheetTableSortConditionArtifact condition, XlsxCellStyleCodec styles)
    {
        element.SetAttributeValue("ref", condition.Reference);
        element.SetAttributeValue("descending", condition.Descending ? "1" : null);
        element.SetAttributeValue("sortBy", null);
        element.SetAttributeValue("iconSet", null);
        element.SetAttributeValue("iconId", null);
        element.SetAttributeValue("dxfId", null);
        element.SetAttributeValue("customList", condition.HasCustomList ? condition.CustomList : null);
        if (condition.Icon is not null)
        {
            element.SetAttributeValue("sortBy", "icon");
            element.SetAttributeValue("iconSet", condition.Icon.IconSet);
            if (condition.Icon.HasIconId) element.SetAttributeValue("iconId", condition.Icon.IconId);
        }
        else if (condition.Color is not null)
        {
            element.SetAttributeValue("sortBy", condition.Color.TargetCase == SpreadsheetTableColorArtifact.TargetOneofCase.CellColor ? "cellColor" : "fontColor");
            element.SetAttributeValue("dxfId", styles.FindOrCreateTableColor(condition.Color, condition.Reference));
        }
    }

    private static bool ValidIcon(SpreadsheetTableIconArtifact icon) =>
        IconSets.TryGetValue(icon.IconSet, out var count) && (!icon.HasIconId || icon.IconId < count);

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

    private static bool TryBool(string? value, bool defaultValue, out bool result)
    {
        if (value is null) { result = defaultValue; return true; }
        if (value is "1" or "true") { result = true; return true; }
        if (value is "0" or "false") { result = false; return true; }
        result = false;
        return false;
    }

    private static bool TryOptionalBool(string? value, out bool? result)
    {
        result = null;
        if (value is null) return true;
        if (!TryBool(value, defaultValue: false, out var parsed)) return false;
        result = parsed;
        return true;
    }

    private static bool TryOptionalUInt(string? value, out uint? result)
    {
        result = null;
        if (value is null) return true;
        if (!uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var parsed)) return false;
        result = parsed;
        return true;
    }

    private static CodecException Invalid(string code, string message, string? location = null) => new(code, message, location);
}
