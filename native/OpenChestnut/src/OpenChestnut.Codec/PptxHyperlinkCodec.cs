using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OpenChestnut.Codec;

internal static class PptxHyperlinkCodec
{
    private const string SlideJumpAction = "ppaction://hlinksldjump";
    private static readonly IReadOnlyDictionary<string, string> ActionUriByName = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        ["nextSlide"] = "ppaction://hlinkshowjump?jump=nextslide",
        ["previousSlide"] = "ppaction://hlinkshowjump?jump=previousslide",
        ["firstSlide"] = "ppaction://hlinkshowjump?jump=firstslide",
        ["lastSlide"] = "ppaction://hlinkshowjump?jump=lastslide",
        ["endShow"] = "ppaction://hlinkshowjump?jump=endshow",
    };
    private static readonly IReadOnlyDictionary<string, string> ActionNameByUri =
        ActionUriByName.ToDictionary(pair => pair.Value, pair => pair.Key, StringComparer.OrdinalIgnoreCase);

    internal static void Read(PresentationTextRun target, A.RunProperties? properties, PptxPartContext? context)
    {
        if (context is not null && TryRead(properties?.GetFirstChild<A.HyperlinkOnClick>(), context, out var hyperlink))
            target.RunHyperlink = hyperlink;
    }

    internal static void Validate(PresentationTextRun run)
    {
        switch (run.HyperlinkCase)
        {
            case PresentationTextRun.HyperlinkOneofCase.None:
                return;
            case PresentationTextRun.HyperlinkOneofCase.NoHyperlink:
                if (!run.NoHyperlink) throw Invalid("Presentation run no_hyperlink choice must be true.");
                return;
            case PresentationTextRun.HyperlinkOneofCase.RunHyperlink:
                Validate(run.RunHyperlink);
                return;
            default:
                throw Invalid("Presentation run hyperlink choice is unsupported.");
        }
    }

    internal static bool HasModeledChoice(PresentationTextRun run) =>
        run.HyperlinkCase != PresentationTextRun.HyperlinkOneofCase.None;

    internal static void Append(A.RunProperties properties, PresentationTextRun source, PptxPartContext? context)
    {
        if (source.HyperlinkCase != PresentationTextRun.HyperlinkOneofCase.RunHyperlink) return;
        if (context is null) throw new CodecException("invalid_presentation_hyperlink", "Presentation hyperlink authoring requires a slide relationship context.");
        properties.Append(Build(source.RunHyperlink, context));
    }

    internal static void Apply(A.RunProperties properties, PresentationTextRun requested, PptxPartContext context)
    {
        var existing = properties.GetFirstChild<A.HyperlinkOnClick>();
        var recognized = TryRead(existing, context, out var current);
        switch (requested.HyperlinkCase)
        {
            case PresentationTextRun.HyperlinkOneofCase.None:
                return;
            case PresentationTextRun.HyperlinkOneofCase.NoHyperlink:
                if (existing is not null && !recognized) throw UnsupportedUnknown();
                existing?.Remove();
                return;
            case PresentationTextRun.HyperlinkOneofCase.RunHyperlink:
                if (existing is not null && !recognized) throw UnsupportedUnknown();
                if (recognized && current.Equals(requested.RunHyperlink)) return;
                var replacement = Build(requested.RunHyperlink, context);
                if (existing is null) properties.Append(replacement);
                else properties.ReplaceChild(replacement, existing);
                return;
            default:
                throw Invalid("Presentation run hyperlink choice is unsupported.");
        }
    }

    internal static void Scrub(A.RunProperties? properties, PptxPartContext? context)
    {
        if (context is null || properties?.GetFirstChild<A.HyperlinkOnClick>() is not { } hyperlink) return;
        if (TryRead(hyperlink, context, out _)) hyperlink.Remove();
    }

    private static bool TryRead(A.HyperlinkOnClick? source, PptxPartContext context, out PresentationRunHyperlink hyperlink)
    {
        hyperlink = new PresentationRunHyperlink();
        if (source is null || source.ChildElements.Count > 0 || source.InvalidUrl is not null || source.EndSound is not null) return false;
        var relationshipId = source.Id?.Value ?? string.Empty;
        var action = source.Action?.Value ?? string.Empty;

        if (relationshipId.Length == 0)
        {
            if (!ActionNameByUri.TryGetValue(action, out var actionName)) return false;
            hyperlink.Action = actionName;
        }
        else if (context.Owner.HyperlinkRelationships.FirstOrDefault(relationship => relationship.Id == relationshipId) is { } external)
        {
            if (!external.IsExternal || action.Length > 0) return false;
            hyperlink.Uri = external.Uri.OriginalString;
        }
        else
        {
            var pair = context.Owner.Parts.FirstOrDefault(candidate => candidate.RelationshipId == relationshipId);
            if (pair.OpenXmlPart is not SlidePart target || !action.Equals(SlideJumpAction, StringComparison.OrdinalIgnoreCase)) return false;
            var path = target.Uri.OriginalString.TrimStart('/');
            if (!context.SlideIdByPartPath.TryGetValue(path, out var slideId)) return false;
            hyperlink.SlideId = slideId;
        }

        if (source.Tooltip is not null) hyperlink.Tooltip = source.Tooltip.Value;
        if (source.TargetFrame is not null) hyperlink.TargetFrame = source.TargetFrame.Value;
        if (source.History is not null) hyperlink.History = source.History.Value;
        if (source.HighlightClick is not null) hyperlink.HighlightClick = source.HighlightClick.Value;
        try
        {
            Validate(hyperlink);
            return true;
        }
        catch (CodecException)
        {
            hyperlink = new PresentationRunHyperlink();
            return false;
        }
    }

    private static A.HyperlinkOnClick Build(PresentationRunHyperlink source, PptxPartContext context)
    {
        Validate(source);
        var hyperlink = new A.HyperlinkOnClick();
        switch (source.TargetCase)
        {
            case PresentationRunHyperlink.TargetOneofCase.Uri:
                hyperlink.Id = context.AddExternalHyperlink(source.Uri);
                break;
            case PresentationRunHyperlink.TargetOneofCase.SlideId:
                hyperlink.Id = context.AddSlide(source.SlideId);
                hyperlink.Action = SlideJumpAction;
                break;
            case PresentationRunHyperlink.TargetOneofCase.Action:
                hyperlink.Id = string.Empty;
                hyperlink.Action = ActionUriByName[source.Action];
                break;
        }
        if (source.HasTooltip) hyperlink.Tooltip = source.Tooltip;
        if (source.HasTargetFrame) hyperlink.TargetFrame = source.TargetFrame;
        if (source.HasHistory) hyperlink.History = source.History;
        if (source.HasHighlightClick) hyperlink.HighlightClick = source.HighlightClick;
        return hyperlink;
    }

    private static void Validate(PresentationRunHyperlink? source)
    {
        if (source is null) throw Invalid("Presentation run hyperlink payload is missing.");
        switch (source.TargetCase)
        {
            case PresentationRunHyperlink.TargetOneofCase.Uri:
                ValidateUri(source.Uri);
                break;
            case PresentationRunHyperlink.TargetOneofCase.SlideId:
                if (string.IsNullOrWhiteSpace(source.SlideId) || source.SlideId.Length > 512)
                    throw Invalid("Presentation run hyperlink slide ID must contain 1 through 512 characters.");
                break;
            case PresentationRunHyperlink.TargetOneofCase.Action:
                if (!ActionUriByName.ContainsKey(source.Action))
                    throw Invalid($"Unsupported Presentation run hyperlink action {source.Action}.");
                break;
            default:
                throw Invalid("Presentation run hyperlink requires exactly one target.");
        }
        if (source.HasTooltip && source.Tooltip.Length > 1024)
            throw Invalid("Presentation run hyperlink tooltip exceeds 1024 characters.");
        if (source.HasTargetFrame && (string.IsNullOrWhiteSpace(source.TargetFrame) || source.TargetFrame.Length > 255))
            throw Invalid("Presentation run hyperlink target frame must contain 1 through 255 characters.");
    }

    private static void ValidateUri(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 4096 || value.Any(character => char.IsControl(character)))
            throw Invalid("Presentation run hyperlink URI must be an absolute URI of at most 4096 characters without controls.");
        if (!Uri.TryCreate(value, UriKind.Absolute, out var uri))
            throw Invalid("Presentation run hyperlink URI must be absolute.");
        if (uri.Scheme.Equals("javascript", StringComparison.OrdinalIgnoreCase) || uri.Scheme.Equals("data", StringComparison.OrdinalIgnoreCase))
            throw Invalid($"Presentation run hyperlink URI uses forbidden scheme {uri.Scheme}.");
    }

    private static CodecException UnsupportedUnknown() => new(
        "unsupported_presentation_edit",
        "Source-preserving PPTX export cannot replace an unmodeled run click action, sound, extension, or malformed relationship.");

    private static CodecException Invalid(string message) => new("invalid_presentation_hyperlink", message);
}
