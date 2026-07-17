using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf.Collections;
using OpenOffice.Artifact.Wire.V1;
using S = DocumentFormat.OpenXml.Spreadsheet;
using TC = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;

namespace OpenChestnut.Codec;

// Owns the small worksheet-native slice used by the Spreadsheet skill:
// ordinary validation rules, four conditional-format profiles, and bounded
// Office 2019 threaded comments with one root plus direct replies. Nested,
// branched, mention-bearing, or otherwise unsupported native collections stay
// hidden behind the validated source package and reject replacement.
internal sealed class XlsxWorksheetFeatureCodec
{
    private const int MaxRulesPerSheet = 4_096;
    private const int MaxCommentsPerSheet = 4_096;
    private const string PersonRelationshipType = "http://schemas.microsoft.com/office/2017/10/relationships/person";
    private const string ThreadedCommentRelationshipType = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";
    private static readonly Regex CellRangePattern = new(
        "^(?<c1>[A-Za-z]{1,3})(?<r1>[1-9][0-9]{0,6})(?::(?<c2>[A-Za-z]{1,3})(?<r2>[1-9][0-9]{0,6}))?$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Regex BracedGuidPattern = new(
        "^\\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\\}$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly HashSet<string> ValidationTypes = new(StringComparer.Ordinal)
    {
        "list", "whole", "decimal", "date", "time", "textLength", "custom",
    };
    private static readonly HashSet<string> ComparisonOperators = new(StringComparer.Ordinal)
    {
        "between", "notBetween", "equal", "notEqual", "lessThan", "lessThanOrEqual", "greaterThan", "greaterThanOrEqual",
    };
    private static readonly HashSet<string> ConditionalTypes = new(StringComparer.Ordinal)
    {
        "cellIs", "expression", "containsText", "colorScale",
    };

    private readonly XlsxCellStyleCodec _styles;
    private readonly HashSet<string> _dirtyPartPaths = new(StringComparer.OrdinalIgnoreCase);

    internal XlsxWorksheetFeatureCodec(XlsxCellStyleCodec styles) => _styles = styles;

    internal IReadOnlySet<string> DirtyPartPaths => _dirtyPartPaths;
    internal bool ThreadedRelationshipGraphDirty { get; private set; }

    internal static bool IsThreadedRelationship(OpaqueOpcRelationship relationship) =>
        relationship.Type.Equals(PersonRelationshipType, StringComparison.Ordinal) ||
        relationship.Type.Equals(ThreadedCommentRelationshipType, StringComparison.Ordinal);

    internal static void Validate(WorksheetArtifact sheet)
    {
        ValidateDataValidations(sheet);
        ValidateConditionalFormats(sheet);
        ValidateThreadedComments(sheet);
    }

    internal void ReadRules(S.Worksheet worksheet, WorksheetArtifact target)
    {
        if (TryReadDataValidations(worksheet, out var validations)) target.DataValidations.Add(validations);
        if (TryReadConditionalFormats(worksheet, out var formats)) target.ConditionalFormats.Add(formats);
    }

    internal void ApplyRules(S.Worksheet worksheet, WorksheetArtifact source, bool sourceBound)
    {
        ApplyDataValidations(worksheet, source, sourceBound);
        ApplyConditionalFormats(worksheet, source, sourceBound);
    }

    internal void ReadThreadedComments(
        WorkbookPart workbookPart,
        IReadOnlyList<(WorksheetPart Part, WorksheetArtifact Artifact)> worksheets)
    {
        var profile = ReadThreadedProfile(workbookPart, worksheets);
        if (!profile.Recognized) return;
        foreach (var (part, artifact) in worksheets)
            if (profile.Comments.TryGetValue(part, out var comments)) artifact.ThreadedComments.Add(comments);
    }

    internal void ApplyThreadedComments(
        WorkbookPart workbookPart,
        IReadOnlyList<(WorksheetPart Part, WorksheetArtifact Artifact)> worksheets,
        bool sourceBound)
    {
        var target = worksheets.ToDictionary(
            item => item.Part,
            item => NormalizeThreadedComments(item.Artifact));
        var nativeIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var comment in target.Values.SelectMany(items => items))
            if (!nativeIds.Add(comment.NativeCommentId))
                throw new CodecException("invalid_spreadsheet_threaded_comment", $"Workbook contains duplicate native threaded-comment id {comment.NativeCommentId}.", comment.CellReference);
        if (!sourceBound)
        {
            if (target.Values.Any(items => items.Count > 0)) RebuildThreadedProfile(workbookPart, worksheets, target);
            return;
        }

        var current = ReadThreadedProfile(workbookPart, worksheets);
        if (!current.Recognized)
        {
            if (target.Values.Any(items => items.Count > 0))
                throw new CodecException("unsupported_spreadsheet_threaded_comment_edit", "Workbook has an unsupported native threaded-comment graph; it cannot be replaced through the bounded profile.", "xl/workbook.xml");
            return;
        }
        if (ThreadedProfilesEqual(current.Comments, target, worksheets)) return;
        RebuildThreadedProfile(workbookPart, worksheets, target);
    }

    private void ApplyDataValidations(S.Worksheet worksheet, WorksheetArtifact source, bool sourceBound)
    {
        var recognized = TryReadDataValidations(worksheet, out var current);
        if (sourceBound && !recognized)
        {
            if (source.DataValidations.Count > 0)
                throw new CodecException("unsupported_spreadsheet_data_validation_edit", $"Worksheet {source.Name} has an unsupported native data-validation collection.", source.Name);
            return;
        }
        if (sourceBound && DataValidationListsEqual(current, source.DataValidations)) return;
        foreach (var collection in worksheet.Elements<S.DataValidations>().ToArray()) collection.Remove();
        if (source.DataValidations.Count == 0) return;

        var container = new S.DataValidations { Count = checked((uint)source.DataValidations.Count) };
        foreach (var item in source.DataValidations)
        {
            var validation = new S.DataValidation
            {
                Type = ValidationType(item.Type),
                AllowBlank = true,
                SequenceOfReferences = References(item.Range),
            };
            if (item.Operator.Length > 0) validation.Operator = ValidationOperator(item.Operator);
            var formula1 = item.Values.Count > 0 ? InlineListFormula(item.Values) : item.Formula1;
            if (formula1.Length > 0) validation.Formula1 = new S.Formula1(formula1);
            if (item.Formula2.Length > 0) validation.Formula2 = new S.Formula2(item.Formula2);
            container.Append(validation);
        }
        InsertBeforeWorksheetTail(worksheet, container);
    }

    private void ApplyConditionalFormats(S.Worksheet worksheet, WorksheetArtifact source, bool sourceBound)
    {
        var recognized = TryReadConditionalFormats(worksheet, out var current);
        if (sourceBound && !recognized)
        {
            if (source.ConditionalFormats.Count > 0)
                throw new CodecException("unsupported_spreadsheet_conditional_format_edit", $"Worksheet {source.Name} has an unsupported native conditional-format collection.", source.Name);
            return;
        }
        if (sourceBound && ConditionalFormatListsEqual(current, source.ConditionalFormats)) return;
        foreach (var collection in worksheet.Elements<S.ConditionalFormatting>().ToArray()) collection.Remove();
        for (var index = 0; index < source.ConditionalFormats.Count; index++)
        {
            var item = source.ConditionalFormats[index];
            var rule = new S.ConditionalFormattingRule
            {
                Type = ConditionalType(item.RuleType),
                Priority = checked((int)(item.Priority > 0 ? item.Priority : checked((uint)index + 1))),
            };
            if (item.Operator.Length > 0) rule.Operator = ConditionalOperator(item.Operator);
            else if (item.RuleType == "containsText") rule.Operator = S.ConditionalFormattingOperatorValues.ContainsText;
            if (item.Text.Length > 0) rule.Text = item.Text;
            if (item.Format is not null) rule.FormatId = _styles.FindOrCreateDifferentialStyle(item.Format, $"{source.Name}!{item.Range}");
            if (item.RuleType == "colorScale") rule.Append(BuildColorScale(item, source.Name));
            else foreach (var formula in item.Formulas) rule.Append(new S.Formula(formula));
            var formatting = new S.ConditionalFormatting { SequenceOfReferences = References(item.Range) };
            formatting.Append(rule);
            InsertBeforeWorksheetTail(worksheet, formatting);
        }
    }

    private bool TryReadDataValidations(S.Worksheet worksheet, out IReadOnlyList<SpreadsheetDataValidationArtifact> artifacts)
    {
        artifacts = [];
        var containers = worksheet.Elements<S.DataValidations>().ToArray();
        if (containers.Length == 0) return true;
        if (containers.Length != 1 || !OnlyAttributes(containers[0], "count") ||
            containers[0].ChildElements.Any(item => item is not S.DataValidation)) return false;
        var validations = containers[0].Elements<S.DataValidation>().ToArray();
        if (containers[0].Count?.HasValue == true && containers[0].Count!.Value != validations.Length) return false;
        var result = new List<SpreadsheetDataValidationArtifact>(validations.Length);
        for (var index = 0; index < validations.Length; index++)
        {
            var item = validations[index];
            if (!OnlyAttributes(item, "type", "operator", "allowBlank", "sqref") || item.AllowBlank?.Value != true ||
                item.ChildElements.Any(child => child is not S.Formula1 and not S.Formula2) ||
                item.Elements<S.Formula1>().Skip(1).Any() || item.Elements<S.Formula2>().Skip(1).Any()) return false;
            var type = ValidationTypeText(item.Type?.Value);
            var operation = ValidationOperatorText(item.Operator?.Value);
            var range = SingleReference(item.SequenceOfReferences);
            if (!ValidationTypes.Contains(type) || range is null || operation.Length > 0 && !ComparisonOperators.Contains(operation)) return false;
            var formula1 = item.Formula1?.Text ?? string.Empty;
            var formula2 = item.Formula2?.Text ?? string.Empty;
            var artifact = new SpreadsheetDataValidationArtifact
            {
                Id = $"data-validation/{index + 1}",
                Range = range,
                Type = type,
                Operator = operation,
                Formula1 = formula1,
                Formula2 = formula2,
            };
            if (type == "list" && TryParseInlineList(formula1, out var values))
            {
                artifact.Formula1 = string.Empty;
                artifact.Values.Add(values);
            }
            try { ValidateDataValidation(artifact, "worksheet"); }
            catch (CodecException) { return false; }
            result.Add(artifact);
        }
        artifacts = result;
        return true;
    }

    private bool TryReadConditionalFormats(S.Worksheet worksheet, out IReadOnlyList<SpreadsheetConditionalFormatArtifact> artifacts)
    {
        artifacts = [];
        var result = new List<SpreadsheetConditionalFormatArtifact>();
        foreach (var formatting in worksheet.Elements<S.ConditionalFormatting>())
        {
            if (!OnlyAttributes(formatting, "sqref") || formatting.ChildElements.Count != 1 ||
                formatting.GetFirstChild<S.ConditionalFormattingRule>() is not { } rule) return false;
            var range = SingleReference(formatting.SequenceOfReferences);
            if (range is null || !OnlyAttributes(rule, "type", "dxfId", "priority", "operator", "text")) return false;
            var ruleType = ConditionalTypeText(rule.Type?.Value);
            if (!ConditionalTypes.Contains(ruleType) || rule.Priority?.HasValue != true || rule.Priority.Value == 0) return false;
            var target = new SpreadsheetConditionalFormatArtifact
            {
                Id = $"conditional-format/{result.Count + 1}",
                Range = range,
                RuleType = ruleType,
                Operator = ConditionalOperatorText(rule.Operator?.Value),
                Text = rule.Text?.Value ?? string.Empty,
                Priority = checked((uint)rule.Priority.Value),
            };
            if (rule.FormatId?.HasValue == true)
            {
                if (!_styles.TryReadDifferentialStyle(rule.FormatId.Value, out var format)) return false;
                target.Format = format;
            }
            if (ruleType == "colorScale")
            {
                if (rule.ChildElements.Count != 1 || rule.GetFirstChild<S.ColorScale>() is not { } colorScale || !TryReadColorScale(colorScale, target)) return false;
            }
            else
            {
                if (rule.ChildElements.Any(item => item is not S.Formula)) return false;
                target.Formulas.Add(rule.Elements<S.Formula>().Select(item => item.Text ?? string.Empty));
            }
            try { ValidateConditionalFormat(target, "worksheet", result.Count); }
            catch (CodecException) { return false; }
            result.Add(target);
        }
        artifacts = result;
        return true;
    }

    private static bool TryReadColorScale(S.ColorScale scale, SpreadsheetConditionalFormatArtifact target)
    {
        if (scale.ChildElements.Any(item => item is not S.ConditionalFormatValueObject and not S.Color)) return false;
        var thresholds = scale.Elements<S.ConditionalFormatValueObject>().ToArray();
        var colors = scale.Elements<S.Color>().ToArray();
        if (thresholds.Length != colors.Length || thresholds.Length is < 2 or > 3) return false;
        if (thresholds[0].Type?.Value != S.ConditionalFormatValueObjectValues.Min ||
            thresholds[^1].Type?.Value != S.ConditionalFormatValueObjectValues.Max) return false;
        if (thresholds.Length == 3 && (thresholds[1].Type?.Value != S.ConditionalFormatValueObjectValues.Percentile || thresholds[1].Val?.Value != "50")) return false;
        if (thresholds.Any(item => !OnlyAttributes(item, "type", "val") || item.HasChildren)) return false;
        foreach (var color in colors)
        {
            if (!OnlyAttributes(color, "rgb", "theme", "indexed", "auto", "tint") || color.HasChildren) return false;
            try
            {
                var value = XlsxCellStyleCodec.ReadConditionalColor(color);
                if (value is null) return false;
                target.Colors.Add(value);
            }
            catch (CodecException) { return false; }
        }
        return true;
    }

    private S.ColorScale BuildColorScale(SpreadsheetConditionalFormatArtifact source, string sheetName)
    {
        var scale = new S.ColorScale();
        scale.Append(new S.ConditionalFormatValueObject { Type = S.ConditionalFormatValueObjectValues.Min });
        if (source.Colors.Count == 3) scale.Append(new S.ConditionalFormatValueObject { Type = S.ConditionalFormatValueObjectValues.Percentile, Val = "50" });
        scale.Append(new S.ConditionalFormatValueObject { Type = S.ConditionalFormatValueObjectValues.Max });
        for (var index = 0; index < source.Colors.Count; index++)
            scale.Append(XlsxCellStyleCodec.WriteConditionalColor(source.Colors[index], $"{sheetName}!{source.Range} color {index + 1}"));
        return scale;
    }

    private static ThreadedProfile ReadThreadedProfile(
        WorkbookPart workbookPart,
        IReadOnlyList<(WorksheetPart Part, WorksheetArtifact Artifact)> worksheets)
    {
        try
        {
            var commentParts = worksheets.ToDictionary(item => item.Part, item => item.Part.WorksheetThreadedCommentsParts.ToArray());
            if (commentParts.Values.Any(parts => parts.Length > 1)) return ThreadedProfile.Unsupported;
            var hasComments = commentParts.Values.Any(parts => parts.Length == 1);
            var personParts = workbookPart.WorkbookPersonParts.ToArray();
            if (!hasComments) return personParts.Length == 0 ? ThreadedProfile.Empty : ThreadedProfile.Unsupported;
            if (personParts.Length != 1 || personParts[0].PersonList is not { } personList ||
                personList.ChildElements.Any(item => item is not TC.Person)) return ThreadedProfile.Unsupported;

            var people = new Dictionary<string, PersonProfile>(StringComparer.Ordinal);
            foreach (var person in personList.Elements<TC.Person>())
            {
                if (!OnlyAttributes(person, "displayName", "id", "userId", "providerId") || person.HasChildren) return ThreadedProfile.Unsupported;
                var id = NormalizeGuid(person.Id?.Value);
                var author = person.DisplayName?.Value ?? string.Empty;
                var userId = person.UserId?.Value ?? string.Empty;
                var providerId = person.ProviderId?.Value ?? string.Empty;
                if (id is null || !ValidText(author, 255, allowLineBreaks: false) || userId.Length > 2_048 || providerId.Length > 255 ||
                    !people.TryAdd(id, new PersonProfile(author, userId, providerId))) return ThreadedProfile.Unsupported;
            }

            var usedPeople = new HashSet<string>(StringComparer.Ordinal);
            var ids = new HashSet<string>(StringComparer.Ordinal);
            var comments = new Dictionary<WorksheetPart, IReadOnlyList<SpreadsheetThreadedCommentArtifact>>(ReferenceEqualityComparer.Instance);
            foreach (var (part, _) in worksheets)
            {
                var nativeParts = commentParts[part];
                if (nativeParts.Length == 0) { comments[part] = []; continue; }
                var root = nativeParts[0].ThreadedComments;
                if (root is null || root.ChildElements.Any(item => item is not TC.ThreadedComment)) return ThreadedProfile.Unsupported;
                var native = new List<NativeThreadedComment>();
                foreach (var item in root.Elements<TC.ThreadedComment>())
                {
                    if (!OnlyAttributes(item, "ref", "dT", "personId", "id", "parentId", "done") ||
                        item.ChildElements.Count != 1 || item.GetFirstChild<TC.ThreadedCommentText>() is not { } text) return ThreadedProfile.Unsupported;
                    var nativeId = NormalizeGuid(item.Id?.Value);
                    var hasParent = item.ParentId?.HasValue == true;
                    var parentId = hasParent ? NormalizeGuid(item.ParentId!.Value) : null;
                    var personId = NormalizeGuid(item.PersonId?.Value);
                    var hasReference = item.Ref?.HasValue == true;
                    var cellReference = hasReference ? CanonicalSingleCell(item.Ref!.Value) : null;
                    if (nativeId is null || hasParent && parentId is null || personId is null || !ids.Add(nativeId) ||
                        !people.TryGetValue(personId, out var person) || !hasParent && cellReference is null || hasReference && cellReference is null)
                        return ThreadedProfile.Unsupported;
                    usedPeople.Add(personId);
                    var dateTime = CanonicalDate(item.DT?.Value);
                    if (dateTime is null || !ValidText(text.Text, 32_767, allowLineBreaks: true)) return ThreadedProfile.Unsupported;
                    native.Add(new NativeThreadedComment(nativeId, parentId, cellReference, personId, text.Text, dateTime, item.Done?.Value ?? false, person));
                }
                var byId = native.ToDictionary(item => item.NativeId, StringComparer.Ordinal);
                var output = new List<SpreadsheetThreadedCommentArtifact>(native.Count);
                foreach (var rootComment in native.Where(item => item.ParentId is null))
                {
                    output.Add(new SpreadsheetThreadedCommentArtifact
                    {
                        Id = $"thread/{rootComment.NativeId}",
                        CellReference = rootComment.CellReference!,
                        NativeCommentId = rootComment.NativeId,
                        Text = rootComment.Text,
                        PersonId = rootComment.PersonId,
                        Author = rootComment.Person.Author,
                        UserId = rootComment.Person.UserId,
                        ProviderId = rootComment.Person.ProviderId,
                        DateTime = rootComment.DateTime,
                        Resolved = rootComment.Resolved,
                    });
                    foreach (var reply in native.Where(item => item.ParentId == rootComment.NativeId))
                    {
                        if (reply.CellReference is not null && !reply.CellReference.Equals(rootComment.CellReference, StringComparison.OrdinalIgnoreCase))
                            return ThreadedProfile.Unsupported;
                        output.Add(new SpreadsheetThreadedCommentArtifact
                        {
                            Id = $"reply/{reply.NativeId}",
                            CellReference = rootComment.CellReference!,
                            NativeCommentId = reply.NativeId,
                            ParentNativeCommentId = rootComment.NativeId,
                            Text = reply.Text,
                            PersonId = reply.PersonId,
                            Author = reply.Person.Author,
                            UserId = reply.Person.UserId,
                            ProviderId = reply.Person.ProviderId,
                            DateTime = reply.DateTime,
                            Resolved = reply.Resolved,
                        });
                    }
                }
                if (output.Count != native.Count || native.Any(item => item.ParentId is not null && (!byId.TryGetValue(item.ParentId, out var parent) || parent.ParentId is not null)))
                    return ThreadedProfile.Unsupported;
                comments[part] = output;
            }
            if (usedPeople.Count != people.Count) return ThreadedProfile.Unsupported;
            return new ThreadedProfile(true, comments);
        }
        catch (Exception exception) when (exception is not OutOfMemoryException)
        {
            return ThreadedProfile.Unsupported;
        }
    }

    private void RebuildThreadedProfile(
        WorkbookPart workbookPart,
        IReadOnlyList<(WorksheetPart Part, WorksheetArtifact Artifact)> worksheets,
        IReadOnlyDictionary<WorksheetPart, IReadOnlyList<SpreadsheetThreadedCommentArtifact>> target)
    {
        var oldPersonParts = workbookPart.WorkbookPersonParts.ToArray();
        foreach (var part in oldPersonParts) _dirtyPartPaths.Add(PartPath(part));
        foreach (var (worksheetPart, _) in worksheets)
        {
            var oldParts = worksheetPart.WorksheetThreadedCommentsParts.ToArray();
            foreach (var part in oldParts) _dirtyPartPaths.Add(PartPath(part));
            if (oldParts.Length > 0 || target[worksheetPart].Count > 0) _dirtyPartPaths.Add(RelationshipPartPath(worksheetPart));
            foreach (var part in oldParts) worksheetPart.DeletePart(part);
        }
        foreach (var part in oldPersonParts) workbookPart.DeletePart(part);

        var allComments = target.Values.SelectMany(items => items).ToArray();
        if (allComments.Length == 0)
        {
            ThreadedRelationshipGraphDirty = oldPersonParts.Length > 0;
            return;
        }
        var people = new Dictionary<string, PersonProfile>(StringComparer.Ordinal);
        foreach (var comment in allComments)
        {
            var profile = new PersonProfile(comment.Author, comment.UserId, comment.ProviderId);
            if (people.TryGetValue(comment.PersonId, out var prior) && prior != profile)
                throw new CodecException("invalid_spreadsheet_threaded_comment", $"Threaded comment person {comment.PersonId} has conflicting metadata.", "xl/workbook.xml");
            people[comment.PersonId] = profile;
        }
        var personPart = workbookPart.AddNewPart<WorkbookPersonPart>();
        var personList = new TC.PersonList();
        foreach (var (id, person) in people.OrderBy(item => item.Key, StringComparer.Ordinal))
        {
            personList.Append(new TC.Person
            {
                Id = id,
                DisplayName = person.Author,
                UserId = person.UserId,
                ProviderId = person.ProviderId,
            });
        }
        personPart.PersonList = personList;
        personList.Save();
        _dirtyPartPaths.Add(PartPath(personPart));

        foreach (var (worksheetPart, _) in worksheets)
        {
            var comments = target[worksheetPart];
            if (comments.Count == 0) continue;
            var part = worksheetPart.AddNewPart<WorksheetThreadedCommentsPart>();
            var root = new TC.ThreadedComments();
            foreach (var comment in comments)
            {
                var native = new TC.ThreadedComment(new TC.ThreadedCommentText(comment.Text))
                {
                    Ref = comment.CellReference,
                    DT = DateTime.Parse(comment.DateTime, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal),
                    PersonId = comment.PersonId,
                    Id = comment.NativeCommentId,
                    Done = comment.Resolved,
                };
                if (comment.ParentNativeCommentId.Length > 0) native.ParentId = comment.ParentNativeCommentId;
                root.Append(native);
            }
            part.ThreadedComments = root;
            root.Save();
            _dirtyPartPaths.Add(PartPath(part));
        }
        ThreadedRelationshipGraphDirty = true;
    }

    private static IReadOnlyList<SpreadsheetThreadedCommentArtifact> NormalizeThreadedComments(WorksheetArtifact sheet)
    {
        var result = new List<SpreadsheetThreadedCommentArtifact>(sheet.ThreadedComments.Count);
        foreach (var source in sheet.ThreadedComments)
        {
            var target = source.Clone();
            target.CellReference = CanonicalSingleCell(source.CellReference) ?? source.CellReference;
            target.NativeCommentId = NormalizeGuid(source.NativeCommentId) ?? DeterministicGuid($"comment:{sheet.Id}:{source.Id}:{target.CellReference}");
            target.Author = string.IsNullOrWhiteSpace(source.Author) ? "User" : source.Author;
            target.UserId = string.IsNullOrEmpty(source.UserId) ? target.Author : source.UserId;
            target.ProviderId = string.IsNullOrEmpty(source.ProviderId) ? "None" : source.ProviderId;
            target.PersonId = NormalizeGuid(source.PersonId) ?? DeterministicGuid($"person:{target.Author}:{target.UserId}:{target.ProviderId}");
            target.DateTime = CanonicalDate(source.DateTime) ?? source.DateTime;
            target.ParentNativeCommentId = source.ParentNativeCommentId.Length == 0
                ? string.Empty
                : NormalizeGuid(source.ParentNativeCommentId) ?? source.ParentNativeCommentId;
            result.Add(target);
        }
        return result;
    }

    private static void ValidateDataValidations(WorksheetArtifact sheet)
    {
        if (sheet.DataValidations.Count > MaxRulesPerSheet) throw InvalidValidation(sheet.Name, "exceeds the 4096-rule budget.");
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var item in sheet.DataValidations)
        {
            if (string.IsNullOrWhiteSpace(item.Id) || !ids.Add(item.Id)) throw InvalidValidation(sheet.Name, "IDs must be non-empty and unique.");
            ValidateDataValidation(item, sheet.Name);
        }
    }

    private static void ValidateDataValidation(SpreadsheetDataValidationArtifact item, string sheetName)
    {
        if (CanonicalRange(item.Range) is null) throw InvalidValidation(sheetName, $"{item.Id} has invalid range {item.Range}.");
        if (!ValidationTypes.Contains(item.Type)) throw InvalidValidation(sheetName, $"{item.Id} has unsupported type {item.Type}.");
        if (item.Operator.Length > 0 && !ComparisonOperators.Contains(item.Operator)) throw InvalidValidation(sheetName, $"{item.Id} has unsupported operator {item.Operator}.");
        if (item.Values.Count > 0 && item.Formula1.Length > 0) throw InvalidValidation(sheetName, $"{item.Id} cannot combine values and formula1.");
        if (item.Values.Count > 0 && item.Type != "list") throw InvalidValidation(sheetName, $"{item.Id} values require list type.");
        if (item.Values.Any(value => value.Contains(',') || !ValidText(value, 255, allowLineBreaks: false)) || item.Values.Count > 256 || InlineListFormula(item.Values).Length > 255)
            throw InvalidValidation(sheetName, $"{item.Id} contains an invalid or oversized inline list.");
        if (!ValidFormula(item.Formula1) || !ValidFormula(item.Formula2)) throw InvalidValidation(sheetName, $"{item.Id} has an invalid formula.");
        if (item.Type == "custom" && item.Formula1.Length == 0 || item.Type == "list" && item.Values.Count == 0 && item.Formula1.Length == 0)
            throw InvalidValidation(sheetName, $"{item.Id} requires formula1 or inline values.");
        var between = item.Operator is "between" or "notBetween";
        if (between && item.Formula2.Length == 0 || !between && item.Formula2.Length > 0)
            throw InvalidValidation(sheetName, $"{item.Id} formula2 is valid only for between/notBetween operators.");
    }

    private static void ValidateConditionalFormats(WorksheetArtifact sheet)
    {
        if (sheet.ConditionalFormats.Count > MaxRulesPerSheet) throw InvalidConditional(sheet.Name, "exceeds the 4096-rule budget.");
        var ids = new HashSet<string>(StringComparer.Ordinal);
        var priorities = new HashSet<uint>();
        for (var index = 0; index < sheet.ConditionalFormats.Count; index++)
        {
            var item = sheet.ConditionalFormats[index];
            if (string.IsNullOrWhiteSpace(item.Id) || !ids.Add(item.Id)) throw InvalidConditional(sheet.Name, "IDs must be non-empty and unique.");
            ValidateConditionalFormat(item, sheet.Name, index);
            var priority = item.Priority > 0 ? item.Priority : checked((uint)index + 1);
            if (!priorities.Add(priority)) throw InvalidConditional(sheet.Name, $"priority {priority} is duplicated.");
        }
    }

    private static void ValidateConditionalFormat(SpreadsheetConditionalFormatArtifact item, string sheetName, int index)
    {
        if (CanonicalRange(item.Range) is null) throw InvalidConditional(sheetName, $"{item.Id} has invalid range {item.Range}.");
        if (!ConditionalTypes.Contains(item.RuleType)) throw InvalidConditional(sheetName, $"{item.Id} has unsupported type {item.RuleType}.");
        if (item.Priority > int.MaxValue) throw InvalidConditional(sheetName, $"{item.Id} priority is too large.");
        if (item.Formulas.Any(formula => !ValidFormula(formula))) throw InvalidConditional(sheetName, $"{item.Id} has an invalid formula.");
        if (item.RuleType == "colorScale")
        {
            if (item.Colors.Count is < 2 or > 3 || item.Formulas.Count > 0 || item.Format is not null || item.Operator.Length > 0 || item.Text.Length > 0)
                throw InvalidConditional(sheetName, $"{item.Id} must contain only two or three color-scale colors.");
            foreach (var color in item.Colors) XlsxCellStyleCodec.WriteConditionalColor(color, $"{sheetName}!{item.Range}");
            return;
        }
        if (item.Colors.Count > 0) throw InvalidConditional(sheetName, $"{item.Id} colors require colorScale type.");
        if (item.RuleType == "expression" && (item.Formulas.Count != 1 || item.Operator.Length > 0 || item.Text.Length > 0))
            throw InvalidConditional(sheetName, $"{item.Id} expression requires exactly one formula.");
        if (item.RuleType == "containsText" && (item.Formulas.Count != 1 || !ValidText(item.Text, 32_767, allowLineBreaks: false) ||
            item.Operator.Length > 0 && item.Operator != "containsText"))
            throw InvalidConditional(sheetName, $"{item.Id} containsText requires text and one formula.");
        if (item.RuleType == "cellIs")
        {
            if (!ComparisonOperators.Contains(item.Operator)) throw InvalidConditional(sheetName, $"{item.Id} has unsupported cellIs operator {item.Operator}.");
            if (item.Text.Length > 0) throw InvalidConditional(sheetName, $"{item.Id} cellIs rules cannot contain text metadata.");
            var expected = item.Operator is "between" or "notBetween" ? 2 : 1;
            if (item.Formulas.Count != expected) throw InvalidConditional(sheetName, $"{item.Id} requires {expected} formula value(s).");
        }
        if (item.Format is not null) XlsxCellStyleCodec.Validate(item.Format, $"{sheetName}!{item.Range}");
    }

    private static void ValidateThreadedComments(WorksheetArtifact sheet)
    {
        if (sheet.ThreadedComments.Count > MaxCommentsPerSheet) throw InvalidComment(sheet.Name, "exceeds the 4096-comment budget.");
        var ids = new HashSet<string>(StringComparer.Ordinal);
        var nativeIds = new HashSet<string>(StringComparer.Ordinal);
        var comments = NormalizeThreadedComments(sheet);
        foreach (var item in comments)
        {
            if (string.IsNullOrWhiteSpace(item.Id) || !ids.Add(item.Id)) throw InvalidComment(sheet.Name, "IDs must be non-empty and unique.");
            if (CanonicalSingleCell(item.CellReference) is null) throw InvalidComment(sheet.Name, $"{item.Id} must target one valid cell.");
            if (!BracedGuidPattern.IsMatch(item.NativeCommentId) || !nativeIds.Add(item.NativeCommentId) || !BracedGuidPattern.IsMatch(item.PersonId))
                throw InvalidComment(sheet.Name, $"{item.Id} has invalid or duplicate native GUID identity.");
            if (item.ParentNativeCommentId.Length > 0 && !BracedGuidPattern.IsMatch(item.ParentNativeCommentId))
                throw InvalidComment(sheet.Name, $"{item.Id} has an invalid parent GUID.");
            if (!ValidText(item.Text, 32_767, allowLineBreaks: true) || !ValidText(item.Author, 255, allowLineBreaks: false) ||
                !ValidText(item.UserId, 2_048, allowLineBreaks: false) || !ValidText(item.ProviderId, 255, allowLineBreaks: false) || CanonicalDate(item.DateTime) is null)
                throw InvalidComment(sheet.Name, $"{item.Id} has invalid text, person metadata, or date-time.");
        }
        var byNativeId = comments.ToDictionary(item => item.NativeCommentId, StringComparer.Ordinal);
        foreach (var item in comments.Where(item => item.ParentNativeCommentId.Length > 0))
        {
            if (!byNativeId.TryGetValue(item.ParentNativeCommentId, out var parent))
                throw InvalidComment(sheet.Name, $"{item.Id} references a missing parent.");
            if (parent.ParentNativeCommentId.Length > 0)
                throw InvalidComment(sheet.Name, $"{item.Id} is nested below another reply; only direct replies are supported.");
            if (!item.CellReference.Equals(parent.CellReference, StringComparison.OrdinalIgnoreCase))
                throw InvalidComment(sheet.Name, $"{item.Id} and its root must target the same cell.");
        }
    }

    private static bool DataValidationListsEqual(IReadOnlyList<SpreadsheetDataValidationArtifact> left, RepeatedField<SpreadsheetDataValidationArtifact> right) =>
        left.Count == right.Count && left.Zip(right).All(pair => DataValidationEqual(pair.First, pair.Second));

    private static bool DataValidationEqual(SpreadsheetDataValidationArtifact left, SpreadsheetDataValidationArtifact right) =>
        left.Range.Equals(right.Range, StringComparison.OrdinalIgnoreCase) && left.Type == right.Type && left.Operator == right.Operator &&
        left.Formula1 == right.Formula1 && left.Formula2 == right.Formula2 && left.Values.SequenceEqual(right.Values, StringComparer.Ordinal);

    private static bool ConditionalFormatListsEqual(IReadOnlyList<SpreadsheetConditionalFormatArtifact> left, RepeatedField<SpreadsheetConditionalFormatArtifact> right) =>
        left.Count == right.Count && left.Zip(right).All(pair => ConditionalFormatEqual(pair.First, pair.Second));

    private static bool ConditionalFormatEqual(SpreadsheetConditionalFormatArtifact left, SpreadsheetConditionalFormatArtifact right) =>
        left.Range.Equals(right.Range, StringComparison.OrdinalIgnoreCase) && left.RuleType == right.RuleType && left.Operator == right.Operator &&
        left.Formulas.SequenceEqual(right.Formulas, StringComparer.Ordinal) && left.Text == right.Text && Equals(left.Format, right.Format) &&
        left.Colors.SequenceEqual(right.Colors) && left.Priority == (right.Priority > 0 ? right.Priority : left.Priority);

    private static bool ThreadedProfilesEqual(
        IReadOnlyDictionary<WorksheetPart, IReadOnlyList<SpreadsheetThreadedCommentArtifact>> left,
        IReadOnlyDictionary<WorksheetPart, IReadOnlyList<SpreadsheetThreadedCommentArtifact>> right,
        IReadOnlyList<(WorksheetPart Part, WorksheetArtifact Artifact)> worksheets) =>
        worksheets.All(item => left.GetValueOrDefault(item.Part, []).Count == right.GetValueOrDefault(item.Part, []).Count &&
            left.GetValueOrDefault(item.Part, []).Zip(right.GetValueOrDefault(item.Part, [])).All(pair => ThreadedCommentEqual(pair.First, pair.Second)));

    private static bool ThreadedCommentEqual(SpreadsheetThreadedCommentArtifact left, SpreadsheetThreadedCommentArtifact right) =>
        left.CellReference.Equals(right.CellReference, StringComparison.OrdinalIgnoreCase) && left.NativeCommentId == right.NativeCommentId &&
        left.Text == right.Text && left.PersonId == right.PersonId && left.Author == right.Author && left.UserId == right.UserId &&
        left.ProviderId == right.ProviderId && CanonicalDate(left.DateTime) == CanonicalDate(right.DateTime) && left.Resolved == right.Resolved &&
        left.ParentNativeCommentId == right.ParentNativeCommentId;

    private static string? SingleReference(ListValue<StringValue>? references)
    {
        var text = references?.InnerText?.Trim() ?? string.Empty;
        return text.Contains(' ') ? null : CanonicalRange(text);
    }

    private static ListValue<StringValue> References(string range) => new() { InnerText = CanonicalRange(range) ?? range };

    private static string? CanonicalSingleCell(string? value)
    {
        var range = CanonicalRange(value);
        return range is not null && !range.Contains(':') ? range : null;
    }

    private static string? CanonicalRange(string? value)
    {
        var match = CellRangePattern.Match(value?.Trim() ?? string.Empty);
        if (!match.Success) return null;
        var column1 = ColumnNumber(match.Groups["c1"].Value);
        var row1 = uint.Parse(match.Groups["r1"].Value, CultureInfo.InvariantCulture);
        var column2 = match.Groups["c2"].Success ? ColumnNumber(match.Groups["c2"].Value) : column1;
        var row2 = match.Groups["r2"].Success ? uint.Parse(match.Groups["r2"].Value, CultureInfo.InvariantCulture) : row1;
        if (column1 is < 1 or > 16_384 || column2 is < 1 or > 16_384 || row1 is < 1 or > 1_048_576 || row2 is < 1 or > 1_048_576 || column2 < column1 || row2 < row1) return null;
        var first = $"{match.Groups["c1"].Value.ToUpperInvariant()}{row1}";
        return match.Groups["c2"].Success ? $"{first}:{match.Groups["c2"].Value.ToUpperInvariant()}{row2}" : first;
    }

    private static int ColumnNumber(string value)
    {
        var number = 0;
        foreach (var character in value) number = checked(number * 26 + char.ToUpperInvariant(character) - 'A' + 1);
        return number;
    }

    private static bool OnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = new HashSet<string>(names, StringComparer.Ordinal);
        return element.GetAttributes().All(attribute => allowed.Contains(attribute.LocalName));
    }

    private static string InlineListFormula(IEnumerable<string> values) => $"\"{string.Join(',', values).Replace("\"", "\"\"")}\"";

    private static bool TryParseInlineList(string value, out IReadOnlyList<string> values)
    {
        values = [];
        if (value.Length < 2 || value[0] != '"' || value[^1] != '"') return false;
        var body = value[1..^1];
        var result = new List<string>();
        var current = new StringBuilder();
        for (var index = 0; index < body.Length; index++)
        {
            if (body[index] == '"' && index + 1 < body.Length && body[index + 1] == '"') { current.Append('"'); index++; continue; }
            if (body[index] == ',') { result.Add(current.ToString()); current.Clear(); continue; }
            current.Append(body[index]);
        }
        result.Add(current.ToString());
        values = result;
        return true;
    }

    private static bool ValidFormula(string value) => value.Length <= 8_192 && !value.Any(character => char.IsControl(character) && character is not '\t' and not '\r' and not '\n');
    private static bool ValidText(string value, int maximum, bool allowLineBreaks) => value.Length > 0 && value.Length <= maximum &&
        !value.Any(character => char.IsControl(character) && (!allowLineBreaks || character is not '\t' and not '\r' and not '\n'));

    private static string? NormalizeGuid(string? value)
    {
        var normalized = value?.Trim().ToUpperInvariant() ?? string.Empty;
        return BracedGuidPattern.IsMatch(normalized) ? normalized : null;
    }

    private static string DeterministicGuid(string seed)
    {
        Span<byte> bytes = stackalloc byte[16];
        SHA256.HashData(Encoding.UTF8.GetBytes(seed)).AsSpan(0, 16).CopyTo(bytes);
        bytes[7] = (byte)((bytes[7] & 0x0F) | 0x40);
        bytes[8] = (byte)((bytes[8] & 0x3F) | 0x80);
        return new Guid(bytes).ToString("B").ToUpperInvariant();
    }

    private static string? CanonicalDate(string? value)
    {
        if (!DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var parsed)) return null;
        return parsed.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'", CultureInfo.InvariantCulture);
    }

    private static string? CanonicalDate(DateTime? value)
    {
        if (value is null) return null;
        var utc = value.Value.Kind == DateTimeKind.Utc ? value.Value : DateTime.SpecifyKind(value.Value, DateTimeKind.Utc);
        return utc.ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'", CultureInfo.InvariantCulture);
    }

    private static S.DataValidationValues ValidationType(string value) => value switch
    {
        "list" => S.DataValidationValues.List,
        "whole" => S.DataValidationValues.Whole,
        "decimal" => S.DataValidationValues.Decimal,
        "date" => S.DataValidationValues.Date,
        "time" => S.DataValidationValues.Time,
        "textLength" => S.DataValidationValues.TextLength,
        "custom" => S.DataValidationValues.Custom,
        _ => throw new InvalidOperationException(),
    };

    private static string ValidationTypeText(S.DataValidationValues? value)
    {
        if (value == S.DataValidationValues.List) return "list";
        if (value == S.DataValidationValues.Whole) return "whole";
        if (value == S.DataValidationValues.Decimal) return "decimal";
        if (value == S.DataValidationValues.Date) return "date";
        if (value == S.DataValidationValues.Time) return "time";
        if (value == S.DataValidationValues.TextLength) return "textLength";
        if (value == S.DataValidationValues.Custom) return "custom";
        return string.Empty;
    }

    private static S.DataValidationOperatorValues ValidationOperator(string value) => value switch
    {
        "between" => S.DataValidationOperatorValues.Between,
        "notBetween" => S.DataValidationOperatorValues.NotBetween,
        "equal" => S.DataValidationOperatorValues.Equal,
        "notEqual" => S.DataValidationOperatorValues.NotEqual,
        "lessThan" => S.DataValidationOperatorValues.LessThan,
        "lessThanOrEqual" => S.DataValidationOperatorValues.LessThanOrEqual,
        "greaterThan" => S.DataValidationOperatorValues.GreaterThan,
        "greaterThanOrEqual" => S.DataValidationOperatorValues.GreaterThanOrEqual,
        _ => throw new InvalidOperationException(),
    };

    private static string ValidationOperatorText(S.DataValidationOperatorValues? value)
    {
        if (value == S.DataValidationOperatorValues.Between) return "between";
        if (value == S.DataValidationOperatorValues.NotBetween) return "notBetween";
        if (value == S.DataValidationOperatorValues.Equal) return "equal";
        if (value == S.DataValidationOperatorValues.NotEqual) return "notEqual";
        if (value == S.DataValidationOperatorValues.LessThan) return "lessThan";
        if (value == S.DataValidationOperatorValues.LessThanOrEqual) return "lessThanOrEqual";
        if (value == S.DataValidationOperatorValues.GreaterThan) return "greaterThan";
        if (value == S.DataValidationOperatorValues.GreaterThanOrEqual) return "greaterThanOrEqual";
        return string.Empty;
    }

    private static S.ConditionalFormatValues ConditionalType(string value) => value switch
    {
        "cellIs" => S.ConditionalFormatValues.CellIs,
        "expression" => S.ConditionalFormatValues.Expression,
        "containsText" => S.ConditionalFormatValues.ContainsText,
        "colorScale" => S.ConditionalFormatValues.ColorScale,
        _ => throw new InvalidOperationException(),
    };

    private static string ConditionalTypeText(S.ConditionalFormatValues? value)
    {
        if (value == S.ConditionalFormatValues.CellIs) return "cellIs";
        if (value == S.ConditionalFormatValues.Expression) return "expression";
        if (value == S.ConditionalFormatValues.ContainsText) return "containsText";
        if (value == S.ConditionalFormatValues.ColorScale) return "colorScale";
        return string.Empty;
    }

    private static S.ConditionalFormattingOperatorValues ConditionalOperator(string value) => value switch
    {
        "between" => S.ConditionalFormattingOperatorValues.Between,
        "notBetween" => S.ConditionalFormattingOperatorValues.NotBetween,
        "equal" => S.ConditionalFormattingOperatorValues.Equal,
        "notEqual" => S.ConditionalFormattingOperatorValues.NotEqual,
        "lessThan" => S.ConditionalFormattingOperatorValues.LessThan,
        "lessThanOrEqual" => S.ConditionalFormattingOperatorValues.LessThanOrEqual,
        "greaterThan" => S.ConditionalFormattingOperatorValues.GreaterThan,
        "greaterThanOrEqual" => S.ConditionalFormattingOperatorValues.GreaterThanOrEqual,
        "containsText" => S.ConditionalFormattingOperatorValues.ContainsText,
        _ => throw new InvalidOperationException(),
    };

    private static string ConditionalOperatorText(S.ConditionalFormattingOperatorValues? value)
    {
        if (value == S.ConditionalFormattingOperatorValues.Between) return "between";
        if (value == S.ConditionalFormattingOperatorValues.NotBetween) return "notBetween";
        if (value == S.ConditionalFormattingOperatorValues.Equal) return "equal";
        if (value == S.ConditionalFormattingOperatorValues.NotEqual) return "notEqual";
        if (value == S.ConditionalFormattingOperatorValues.LessThan) return "lessThan";
        if (value == S.ConditionalFormattingOperatorValues.LessThanOrEqual) return "lessThanOrEqual";
        if (value == S.ConditionalFormattingOperatorValues.GreaterThan) return "greaterThan";
        if (value == S.ConditionalFormattingOperatorValues.GreaterThanOrEqual) return "greaterThanOrEqual";
        if (value == S.ConditionalFormattingOperatorValues.ContainsText) return "containsText";
        return string.Empty;
    }

    private static void InsertBeforeWorksheetTail(S.Worksheet worksheet, OpenXmlElement element)
    {
        var before = worksheet.ChildElements.FirstOrDefault(item => item is S.DataValidations or S.Hyperlinks or S.PrintOptions or S.PageMargins or S.PageSetup or S.HeaderFooter or S.RowBreaks or S.ColumnBreaks or S.CustomProperties or S.CellWatches or S.IgnoredErrors or S.Drawing or S.LegacyDrawing or S.LegacyDrawingHeaderFooter or S.Picture or S.OleObjects or S.Controls or S.WebPublishItems or S.TableParts or S.WorksheetExtensionList);
        if (element is S.DataValidations) before = worksheet.ChildElements.FirstOrDefault(item => item is S.Hyperlinks or S.PrintOptions or S.PageMargins or S.PageSetup or S.HeaderFooter or S.RowBreaks or S.ColumnBreaks or S.CustomProperties or S.CellWatches or S.IgnoredErrors or S.Drawing or S.LegacyDrawing or S.LegacyDrawingHeaderFooter or S.Picture or S.OleObjects or S.Controls or S.WebPublishItems or S.TableParts or S.WorksheetExtensionList);
        if (before is null) worksheet.Append(element);
        else worksheet.InsertBefore(element, before);
    }

    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');

    private static string RelationshipPartPath(OpenXmlPart part)
    {
        var path = PartPath(part);
        var slash = path.LastIndexOf('/');
        var directory = slash < 0 ? string.Empty : path[..slash];
        var file = slash < 0 ? path : path[(slash + 1)..];
        return directory.Length == 0 ? $"_rels/{file}.rels" : $"{directory}/_rels/{file}.rels";
    }

    private static CodecException InvalidValidation(string sheetName, string message) => new("invalid_spreadsheet_data_validation", $"Worksheet {sheetName} data validation {message}", sheetName);
    private static CodecException InvalidConditional(string sheetName, string message) => new("invalid_spreadsheet_conditional_format", $"Worksheet {sheetName} conditional format {message}", sheetName);
    private static CodecException InvalidComment(string sheetName, string message) => new("invalid_spreadsheet_threaded_comment", $"Worksheet {sheetName} threaded comment {message}", sheetName);

    private sealed record PersonProfile(string Author, string UserId, string ProviderId);
    private sealed record NativeThreadedComment(
        string NativeId,
        string? ParentId,
        string? CellReference,
        string PersonId,
        string Text,
        string DateTime,
        bool Resolved,
        PersonProfile Person);
    private sealed record ThreadedProfile(bool Recognized, IReadOnlyDictionary<WorksheetPart, IReadOnlyList<SpreadsheetThreadedCommentArtifact>> Comments)
    {
        internal static ThreadedProfile Empty { get; } = new(true, new Dictionary<WorksheetPart, IReadOnlyList<SpreadsheetThreadedCommentArtifact>>(ReferenceEqualityComparer.Instance));
        internal static ThreadedProfile Unsupported { get; } = new(false, new Dictionary<WorksheetPart, IReadOnlyList<SpreadsheetThreadedCommentArtifact>>(ReferenceEqualityComparer.Instance));
    }
}
