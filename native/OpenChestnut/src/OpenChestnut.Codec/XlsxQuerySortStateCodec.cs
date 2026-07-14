using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns the recognized refresh-local sortState branch of a QueryTablePart.
// Unlike ordinary table authoring, this source-bound codec patches the existing
// condition elements in place so extLst and all surrounding refresh XML remain
// untouched. Presence and condition count are source topology.
internal sealed class XlsxQuerySortStateCodec
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

    private readonly XElement _element;
    private readonly XlsxCellStyleCodec _styles;
    private readonly (uint Top, uint Left, uint Bottom, uint Right) _tableBounds;
    private readonly SpreadsheetTableSortStateArtifact _source;

    private XlsxQuerySortStateCodec(
        XElement element,
        XlsxCellStyleCodec styles,
        (uint Top, uint Left, uint Bottom, uint Right) tableBounds,
        SpreadsheetTableSortStateArtifact artifact)
    {
        _element = element;
        _styles = styles;
        _tableBounds = tableBounds;
        _source = artifact.Clone();
        Artifact = artifact;
    }

    internal SpreadsheetTableSortStateArtifact Artifact { get; private set; }

    internal static bool TryCreate(
        XElement element,
        XlsxCellStyleCodec styles,
        (uint Top, uint Left, uint Bottom, uint Right) tableBounds,
        out XlsxQuerySortStateCodec? codec)
    {
        codec = null;
        if (!TryRead(element, styles, out var artifact)) return false;
        try
        {
            Validate(artifact!, tableBounds, string.Empty);
        }
        catch (CodecException)
        {
            return false;
        }
        codec = new XlsxQuerySortStateCodec(element, styles, tableBounds, artifact!);
        return true;
    }

    internal void ValidateShape(SpreadsheetTableSortStateArtifact? desired, string location)
    {
        if (desired is null)
            throw Invalid("Source-preserving XLSX export cannot remove a recognized query refresh sort state.", location);
        if (desired.Conditions.Count != _source.Conditions.Count)
            throw Invalid("Source-preserving XLSX export cannot add or remove query refresh sort conditions.", location);
        Validate(desired, _tableBounds, location);
    }

    internal void Patch(SpreadsheetTableSortStateArtifact desired, string location)
    {
        ValidateShape(desired, location);
        _element.SetAttributeValue("ref", desired.Reference);
        _element.SetAttributeValue("caseSensitive", desired.CaseSensitive ? "1" : null);
        _element.SetAttributeValue("sortMethod", desired.HasSortMethod ? desired.SortMethod : null);
        var conditions = _element.Elements(Spreadsheet + "sortCondition").ToArray();
        for (var index = 0; index < conditions.Length; index++)
        {
            var element = conditions[index];
            var condition = desired.Conditions[index];
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
                element.SetAttributeValue("dxfId", _styles.FindOrCreateTableColor(condition.Color, condition.Reference));
            }
        }
        Artifact = desired.Clone();
    }

    internal static IEnumerable<string> Semantics(SpreadsheetTableSortStateArtifact? sort)
    {
        if (sort is null) { yield return "no-query-sort-state"; yield break; }
        yield return sort.Reference;
        yield return sort.CaseSensitive.ToString();
        yield return sort.HasSortMethod ? sort.SortMethod : "no-sort-method";
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

    private static bool TryRead(XElement element, XlsxCellStyleCodec styles, out SpreadsheetTableSortStateArtifact? artifact)
    {
        artifact = null;
        if (element.Name != Spreadsheet + "sortState" || element.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration &&
            (attribute.Name.Namespace != XNamespace.None || attribute.Name.LocalName is not ("ref" or "caseSensitive" or "sortMethod"))) ||
            element.Attribute("ref") is not XAttribute reference ||
            !TryBool(element.Attribute("caseSensitive")?.Value, defaultValue: false, out var caseSensitive)) return false;
        var sortMethod = element.Attribute("sortMethod")?.Value;
        if (sortMethod is not null && !SortMethods.Contains(sortMethod)) return false;
        var children = element.Elements().ToArray();
        var extensionIndexes = children.Select((child, index) => (child, index)).Where(item => item.child.Name == Spreadsheet + "extLst").Select(item => item.index).ToArray();
        if (extensionIndexes.Length > 1 || extensionIndexes.Length == 1 && extensionIndexes[0] != children.Length - 1 ||
            children.Any(child => child.Name != Spreadsheet + "sortCondition" && child.Name != Spreadsheet + "extLst")) return false;
        var conditions = children.Where(child => child.Name == Spreadsheet + "sortCondition").ToArray();
        if (conditions.Length is < 1 or > 64) return false;
        var sort = new SpreadsheetTableSortStateArtifact { Reference = reference.Value, CaseSensitive = caseSensitive };
        if (sortMethod is not null) sort.SortMethod = sortMethod;
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

    private static void Validate(
        SpreadsheetTableSortStateArtifact sort,
        (uint Top, uint Left, uint Bottom, uint Right) tableBounds,
        string location)
    {
        if (!TryRange(sort.Reference, out var bounds) || bounds.Top < tableBounds.Top || bounds.Left < tableBounds.Left ||
            bounds.Bottom > tableBounds.Bottom || bounds.Right > tableBounds.Right)
            throw Invalid("Worksheet query refresh sort range must be contained in the source table range.", location);
        if (sort.Conditions.Count is < 1 or > 64)
            throw Invalid("Worksheet query refresh sort state must provide between one and 64 conditions.", location);
        if (sort.HasSortMethod && !SortMethods.Contains(sort.SortMethod))
            throw Invalid("Worksheet query refresh has an invalid locale-specific sort method.", location);
        var columns = new HashSet<uint>();
        foreach (var condition in sort.Conditions)
        {
            if (!TryRange(condition.Reference, out var conditionBounds) || conditionBounds.Left != conditionBounds.Right ||
                conditionBounds.Top != bounds.Top || conditionBounds.Bottom != bounds.Bottom || conditionBounds.Left < bounds.Left ||
                conditionBounds.Right > bounds.Right || !columns.Add(conditionBounds.Left))
                throw Invalid("Worksheet query refresh has an invalid or duplicate sort condition.", location);
            if (condition.Icon is not null && !ValidIcon(condition.Icon))
                throw Invalid("Worksheet query refresh has an invalid icon-sort condition.", location);
            if (condition.Icon is not null && condition.Color is not null)
                throw Invalid("Worksheet query refresh sort condition cannot combine icon and color selectors.", location);
            if (condition.HasCustomList && (string.IsNullOrWhiteSpace(condition.CustomList) || condition.CustomList.Length > 32_767 ||
                condition.CustomList.Any(char.IsControl) || condition.Icon is not null || condition.Color is not null))
                throw Invalid("Worksheet query refresh has an invalid custom-list value-sort condition.", location);
            if (condition.Color is not null) XlsxCellStyleCodec.ValidateTableColor(condition.Color, "query refresh sort");
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

    private static bool TryOptionalUInt(string? value, out uint? result)
    {
        result = null;
        if (value is null) return true;
        if (!uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var parsed)) return false;
        result = parsed;
        return true;
    }

    private static CodecException Invalid(string message, string? location = null) => new("invalid_worksheet_table", message, location);
}
