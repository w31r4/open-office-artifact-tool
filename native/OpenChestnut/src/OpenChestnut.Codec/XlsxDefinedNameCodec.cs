using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns the bounded workbook.xml definedNames projection. Ordinary range names
// are editable; reserved built-ins, formulas/constants, macro attributes, and
// other advanced profiles remain hidden in immutable source slots.
internal sealed class XlsxDefinedNameCodec
{
    private const int MaxDefinedNames = 65_536;
    private const int MaxNameLength = 255;
    private const int MaxCommentLength = 255;
    private const int MaxFormulaLength = 8_192;
    private static readonly Regex NamePattern = new(@"^[A-Za-z_\\][A-Za-z0-9_.\\]*$", RegexOptions.CultureInvariant);
    private static readonly Regex R1C1Pattern = new(@"^(?:R(?:[1-9][0-9]*)?C(?:[1-9][0-9]*)?|R|C)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex RangePattern = new(@"^(?:=)?(?<sheet>'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!(?<range>\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})?)$", RegexOptions.CultureInvariant);
    private readonly Workbook _workbook;
    private readonly string[] _sourceSheetNames;
    private readonly List<Entry> _entries = [];

    private sealed class Entry
    {
        internal required DefinedName Element { get; init; }
        internal required uint Ordinal { get; init; }
        internal required string Name { get; init; }
        internal uint? LocalSheetId { get; init; }
        internal SpreadsheetDefinedNameArtifact? SourceArtifact { get; init; }
    }

    internal XlsxDefinedNameCodec(WorkbookPart workbookPart, IReadOnlyList<string> sourceSheetNames)
    {
        _workbook = workbookPart.Workbook ?? throw Invalid("Workbook root is missing.");
        _sourceSheetNames = sourceSheetNames.ToArray();
        var container = _workbook.DefinedNames;
        if (container is null)
        {
            WorkbookXmlSha256 = string.Empty;
            IsReadable = true;
            return;
        }
        WorkbookXmlSha256 = PartSha256(workbookPart);
        if (container.ChildElements.Count > MaxDefinedNames || container.ChildElements.Any(child => child is not DefinedName)) return;
        var ordinal = 0U;
        foreach (var element in container.Elements<DefinedName>())
        {
            var name = element.Name?.Value ?? string.Empty;
            var localSheetId = element.LocalSheetId?.Value;
            SpreadsheetDefinedNameArtifact? artifact = null;
            if (TryRead(element, ordinal, _sourceSheetNames, out var recognized))
            {
                recognized!.Source = new SpreadsheetDefinedNameSourceBinding
                {
                    Ordinal = ordinal,
                    WorkbookXmlSha256 = WorkbookXmlSha256,
                    DefinedNameXmlSha256 = ElementSha256(element),
                    SemanticSha256 = SemanticSha256(recognized),
                    Editable = true,
                };
                artifact = recognized;
            }
            _entries.Add(new Entry { Element = element, Ordinal = ordinal, Name = name, LocalSheetId = localSheetId, SourceArtifact = artifact });
            ordinal++;
        }
        IsReadable = true;
    }

    internal bool IsReadable { get; }
    internal string WorkbookXmlSha256 { get; }
    internal IReadOnlyList<SpreadsheetDefinedNameArtifact> Read() => _entries
        .Where(entry => entry.SourceArtifact is not null)
        .Select(entry => entry.SourceArtifact!.Clone())
        .ToArray();

    internal void Apply(IReadOnlyList<SpreadsheetDefinedNameArtifact> desired, bool sourceBound, IReadOnlyList<string> targetSheetNames)
    {
        ValidateArtifact(desired, targetSheetNames);
        if (!sourceBound)
        {
            if (desired.Count == 0) return;
            var container = new DefinedNames(desired.Select(item => Create(item, targetSheetNames)));
            _workbook.DefinedNames = container;
            return;
        }
        if (!IsReadable)
        {
            if (desired.Count > 0) throw Invalid("Source-preserving XLSX export cannot replace an opaque definedNames container.");
            return;
        }

        var recognized = _entries.Where(entry => entry.SourceArtifact is not null).ToArray();
        if (desired.Count != recognized.Length)
            throw Invalid("Source-preserving XLSX export cannot add or remove recognized defined names.");
        ValidateAgainstOpaqueNames(desired, targetSheetNames);
        for (var index = 0; index < recognized.Length; index++)
        {
            var entry = recognized[index];
            var source = entry.SourceArtifact!;
            var target = desired[index];
            ValidateBinding(target.Source, source.Source, entry.Ordinal);
            if (SemanticSha256(target).Equals(source.Source.SemanticSha256, StringComparison.OrdinalIgnoreCase)) continue;
            Patch(entry.Element, target, targetSheetNames);
        }
    }

    internal static void ValidateArtifact(IReadOnlyList<SpreadsheetDefinedNameArtifact> names, IReadOnlyList<string> sheetNames)
    {
        if (names.Count > MaxDefinedNames) throw Invalid($"Workbook exceeds the bounded {MaxDefinedNames} defined-name limit.");
        var unique = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in names)
        {
            Validate(item, sheetNames);
            var scope = ScopeIndex(item, sheetNames);
            if (!unique.Add(Key(item.Name, scope)))
                throw Invalid($"Defined name {item.Name} is duplicated in the same workbook or worksheet scope.");
        }
    }

    private void ValidateAgainstOpaqueNames(IReadOnlyList<SpreadsheetDefinedNameArtifact> desired, IReadOnlyList<string> targetSheetNames)
    {
        var unique = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in _entries.Where(entry => entry.SourceArtifact is null))
            if (!string.IsNullOrWhiteSpace(entry.Name) && !unique.Add(Key(entry.Name, entry.LocalSheetId)))
                throw Invalid("Source workbook has ambiguous opaque defined-name identity.");
        foreach (var item in desired)
            if (!unique.Add(Key(item.Name, ScopeIndex(item, targetSheetNames))))
                throw Invalid($"Defined name {item.Name} collides with an opaque source-defined name.");
    }

    private static bool TryRead(DefinedName element, uint ordinal, IReadOnlyList<string> sheetNames, out SpreadsheetDefinedNameArtifact? artifact)
    {
        artifact = null;
        if (element.Function is not null || element.VbProcedure is not null || element.Xlm is not null ||
            element.FunctionGroupId is not null || element.ShortcutKey is not null || element.PublishToServer is not null ||
            element.WorkbookParameter is not null || element.CustomMenu is not null || element.Description is not null ||
            element.Help is not null || element.StatusBar is not null || element.ExtendedAttributes.Any()) return false;
        var result = new SpreadsheetDefinedNameArtifact
        {
            Id = $"defined-name/{ordinal + 1}",
            Name = element.Name?.Value ?? string.Empty,
            RefersTo = element.Text ?? string.Empty,
        };
        if (element.LocalSheetId?.Value is uint localSheetId)
        {
            if (localSheetId >= sheetNames.Count) return false;
            result.ScopeSheetName = sheetNames[(int)localSheetId];
        }
        if (element.Comment is not null) result.Comment = element.Comment.Value;
        if (element.Hidden is not null) result.Hidden = element.Hidden.Value;
        try
        {
            Validate(result, sheetNames, rejectReserved: true);
            artifact = result;
            return true;
        }
        catch (CodecException)
        {
            return false;
        }
    }

    private static void Validate(SpreadsheetDefinedNameArtifact item, IReadOnlyList<string> sheetNames, bool rejectReserved = true)
    {
        if (item.Id.Length > MaxNameLength || HasInvalidXmlControl(item.Id) ||
            item.Name.Length is < 1 or > MaxNameLength || HasInvalidXmlControl(item.Name) || !NamePattern.IsMatch(item.Name) ||
            R1C1Pattern.IsMatch(item.Name) || IsA1Reference(item.Name) ||
            rejectReserved && item.Name.StartsWith("_xl", StringComparison.OrdinalIgnoreCase))
            throw Invalid($"Defined name {item.Name} is outside the bounded public name profile.");
        if (item.RefersTo.Length is < 1 or > MaxFormulaLength || HasInvalidXmlControl(item.RefersTo) || !TryRangeSheet(item.RefersTo, out var formulaSheet) ||
            !sheetNames.Contains(formulaSheet, StringComparer.OrdinalIgnoreCase))
            throw Invalid($"Defined name {item.Name} must refer to a bounded range on an existing worksheet.");
        if (item.HasScopeSheetName && (string.IsNullOrWhiteSpace(item.ScopeSheetName) || !sheetNames.Contains(item.ScopeSheetName, StringComparer.OrdinalIgnoreCase)))
            throw Invalid($"Defined name {item.Name} has an unknown worksheet scope {item.ScopeSheetName}.");
        if (item.HasComment && (item.Comment.Length > MaxCommentLength || HasInvalidXmlControl(item.Comment)))
            throw Invalid($"Defined name {item.Name} comment exceeds the bounded Excel profile.");
    }

    private static DefinedName Create(SpreadsheetDefinedNameArtifact item, IReadOnlyList<string> sheetNames)
    {
        var result = new DefinedName(item.RefersTo) { Name = item.Name };
        Patch(result, item, sheetNames);
        return result;
    }

    private static void Patch(DefinedName element, SpreadsheetDefinedNameArtifact item, IReadOnlyList<string> sheetNames)
    {
        element.Name = item.Name;
        element.Text = item.RefersTo;
        if (item.HasScopeSheetName) element.LocalSheetId = checked((uint)ScopeIndex(item, sheetNames)!.Value);
        else element.LocalSheetId = null;
        if (item.HasComment) element.Comment = item.Comment;
        else element.Comment = null;
        if (item.HasHidden) element.Hidden = item.Hidden;
        else element.Hidden = null;
    }

    private static void ValidateBinding(SpreadsheetDefinedNameSourceBinding? desired, SpreadsheetDefinedNameSourceBinding source, uint ordinal)
    {
        if (desired is null || desired.Ordinal != ordinal ||
            !desired.WorkbookXmlSha256.Equals(source.WorkbookXmlSha256, StringComparison.OrdinalIgnoreCase) ||
            !desired.DefinedNameXmlSha256.Equals(source.DefinedNameXmlSha256, StringComparison.OrdinalIgnoreCase) ||
            !desired.SemanticSha256.Equals(source.SemanticSha256, StringComparison.OrdinalIgnoreCase) ||
            desired.Editable != source.Editable)
            throw Invalid($"Defined name source binding at ordinal {ordinal} does not match the validated source workbook.");
    }

    private static uint? ScopeIndex(SpreadsheetDefinedNameArtifact item, IReadOnlyList<string> sheetNames)
    {
        if (!item.HasScopeSheetName) return null;
        for (var index = 0; index < sheetNames.Count; index++)
            if (sheetNames[index].Equals(item.ScopeSheetName, StringComparison.OrdinalIgnoreCase)) return checked((uint)index);
        return null;
    }

    private static bool TryRangeSheet(string formula, out string sheetName)
    {
        sheetName = string.Empty;
        var match = RangePattern.Match(formula);
        if (!match.Success) return false;
        var token = match.Groups["sheet"].Value;
        sheetName = token.StartsWith("'", StringComparison.Ordinal) ? token[1..^1].Replace("''", "'", StringComparison.Ordinal) : token;
        return match.Groups["range"].Value.Split(':').All(IsA1Reference);
    }

    private static bool IsA1Reference(string value)
    {
        var token = value.Replace("$", string.Empty, StringComparison.Ordinal);
        var letters = new string(token.TakeWhile(char.IsLetter).ToArray());
        if (letters.Length is < 1 or > 3 || !uint.TryParse(token[letters.Length..], out var row) || row is < 1 or > 1_048_576) return false;
        var column = 0U;
        foreach (var character in letters.ToUpperInvariant()) column = checked(column * 26 + (uint)(character - 'A' + 1));
        return column is >= 1 and <= 16_384;
    }

    private static bool HasInvalidXmlControl(string value) => value.Any(character => char.IsControl(character) && character is not '\t' and not '\n' and not '\r');
    private static string Key(string name, uint? scope) => $"{(scope.HasValue ? scope.Value.ToString() : "workbook")}\0{name}";
    private static string SemanticSha256(SpreadsheetDefinedNameArtifact item) => Sha256(Encoding.UTF8.GetBytes(string.Join("\0",
        item.Id, item.Name, item.RefersTo,
        item.HasScopeSheetName ? $"scope:{item.ScopeSheetName}" : "scope:-",
        item.HasComment ? $"comment:{item.Comment}" : "comment:-",
        item.HasHidden ? $"hidden:{item.Hidden}" : "hidden:-")));
    private static string ElementSha256(DefinedName element) => Sha256(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string PartSha256(WorkbookPart part)
    {
        using var source = part.GetStream(FileMode.Open, FileAccess.Read);
        return Convert.ToHexString(SHA256.HashData(source)).ToLowerInvariant();
    }
    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static CodecException Invalid(string message) => new("invalid_workbook_defined_name", message, "xl/workbook.xml");
}
