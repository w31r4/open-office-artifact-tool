using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenChestnut.Codec;

internal sealed record PptxCustomShowReadResult(
    IReadOnlyList<PresentationCustomShowArtifact> Shows,
    bool Opaque,
    ulong SemanticItems,
    string? Reason);

// Owns the bounded inline p:custShowLst graph. Canonical lists are semantic;
// every other list stays byte-preserved in the validated source package and
// is never reconstructed from an incomplete model.
internal static class PptxCustomShowCodec
{
    internal const int MaxCustomShows = 4_096;
    internal const int MaxSlidesPerShow = 16_384;
    private const int MaxNameLength = 255;

    internal static PptxCustomShowReadResult Read(
        PresentationPart owner,
        IReadOnlyDictionary<string, string> slideIdByRelationshipId,
        EffectiveCodecLimits limits)
    {
        var root = owner.Presentation ??
            throw new CodecException("missing_presentation_root", "PPTX package has no Presentation root.", "ppt/presentation.xml");
        var list = root.CustomShowList;
        if (list is null) return new PptxCustomShowReadResult([], false, 0, null);

        var childCount = list.ChildElements.Count;
        if (childCount > MaxCustomShows)
            throw new CodecException(
                "presentation_custom_show_budget_exceeded",
                $"PPTX custom-show list has {childCount} children and exceeds the {MaxCustomShows}-show budget.",
                "ppt/presentation.xml");
        if (childCount == 0)
            return Opaque("the custom-show list is empty", 0);
        if (HasExtendedAttributes(list) || list.ChildElements.Any(child => child is not P.CustomShow))
            return Opaque("the custom-show list contains extension attributes or non-show children", checked((ulong)childCount));

        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var nativeIds = new HashSet<uint>();
        var shows = new List<PresentationCustomShowArtifact>(childCount);
        ulong semanticItems = 0;
        var sourceShows = list.Elements<P.CustomShow>().ToArray();
        for (var showIndex = 0; showIndex < sourceShows.Length; showIndex++)
        {
            var source = sourceShows[showIndex];
            semanticItems = checked(semanticItems + 1);
            if (semanticItems > limits.MaxCells)
                throw new CodecException(
                    "presentation_item_budget_exceeded",
                    $"PPTX custom shows exceed max_cells semantic-item budget ({limits.MaxCells}).",
                    "ppt/presentation.xml");
            if (HasExtendedAttributes(source) || source.ChildElements.Count != 1 || source.ChildElements[0] is not P.SlideList slideList)
                return Opaque($"custom show {showIndex + 1} is not one canonical slide-list child", semanticItems);

            var name = source.Name?.Value ?? string.Empty;
            if (!ValidName(name) || !names.Add(name))
                return Opaque($"custom show {showIndex + 1} has a missing, invalid, or duplicate name", semanticItems);
            if (source.Id?.Value is not { } nativeId || !nativeIds.Add(nativeId))
                return Opaque($"custom show {showIndex + 1} has a missing or duplicate native ID", semanticItems);
            if (HasExtendedAttributes(slideList) || slideList.ChildElements.Any(child => child is not P.SlideListEntry))
                return Opaque($"custom show {showIndex + 1} contains extension attributes or non-slide children", semanticItems);

            var entries = slideList.Elements<P.SlideListEntry>().ToArray();
            if (entries.Length == 0)
                return Opaque($"custom show {showIndex + 1} has no slides", semanticItems);
            if (entries.Length > MaxSlidesPerShow)
                throw new CodecException(
                    "presentation_custom_show_budget_exceeded",
                    $"PPTX custom show {showIndex + 1} has {entries.Length} slides and exceeds the {MaxSlidesPerShow}-slide budget.",
                    "ppt/presentation.xml");
            semanticItems = checked(semanticItems + (ulong)entries.Length);
            if (semanticItems > limits.MaxCells)
                throw new CodecException(
                    "presentation_item_budget_exceeded",
                    $"PPTX custom shows exceed max_cells semantic-item budget ({limits.MaxCells}).",
                    "ppt/presentation.xml");

            var show = new PresentationCustomShowArtifact
            {
                Id = $"custom-show/{showIndex + 1}",
                Name = name,
                NativeId = nativeId,
            };
            foreach (var entry in entries)
            {
                if (HasExtendedAttributes(entry))
                    return Opaque($"custom show {showIndex + 1} contains an extended slide reference", semanticItems);
                var relationshipId = entry.Id?.Value ?? string.Empty;
                if (relationshipId.Length == 0 || !slideIdByRelationshipId.TryGetValue(relationshipId, out var slideId))
                    return Opaque($"custom show {showIndex + 1} contains an unresolved or out-of-presentation slide relationship", semanticItems);
                show.SlideIds.Add(slideId);
            }
            show.Source = new PresentationCustomShowSourceBinding
            {
                Ordinal = checked((uint)showIndex),
                ShowXmlSha256 = HashElement(source),
                Editable = true,
            };
            show.Source.SemanticSha256 = SemanticHash(show);
            shows.Add(show);
        }
        return new PptxCustomShowReadResult(shows, false, semanticItems, null);
    }

    internal static void Validate(
        PresentationArtifact artifact,
        bool hasSourcePackage,
        EffectiveCodecLimits limits,
        ref ulong semanticItems)
    {
        if (artifact.CustomShowsOpaque)
        {
            if (!hasSourcePackage || artifact.CustomShows.Count != 0)
                throw new CodecException(
                    "invalid_presentation_custom_show",
                    "Opaque custom shows require a validated source package and cannot coexist with semantic custom shows.");
            return;
        }
        if (artifact.CustomShows.Count > MaxCustomShows)
            throw new CodecException(
                "presentation_custom_show_budget_exceeded",
                $"Presentation cannot contain more than {MaxCustomShows} custom shows.");

        var slideIds = artifact.Slides.Select(slide => slide.Id).ToHashSet(StringComparer.Ordinal);
        if (artifact.CustomShows.Count > 0 && (slideIds.Count != artifact.Slides.Count || slideIds.Contains(string.Empty)))
            throw new CodecException("invalid_presentation_slide", "Presentation slide IDs must be non-empty and unique when custom shows are present.");
        var ids = new HashSet<string>(StringComparer.Ordinal);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var nativeIds = new HashSet<uint>();
        for (var showIndex = 0; showIndex < artifact.CustomShows.Count; showIndex++)
        {
            var show = artifact.CustomShows[showIndex];
            semanticItems = checked(semanticItems + 1 + (ulong)show.SlideIds.Count);
            if (semanticItems > limits.MaxCells)
                throw new CodecException(
                    "presentation_item_budget_exceeded",
                    $"Presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).");
            if (string.IsNullOrWhiteSpace(show.Id) || show.Id.Length > 1_024 || !ids.Add(show.Id))
                throw new CodecException("invalid_presentation_custom_show", "Presentation custom-show IDs must be non-empty, bounded, and unique.");
            if (!ValidName(show.Name) || !names.Add(show.Name))
                throw new CodecException("invalid_presentation_custom_show", "Presentation custom-show names must be trimmed, valid, and case-insensitively unique.");
            if (!nativeIds.Add(show.NativeId))
                throw new CodecException("invalid_presentation_custom_show", "Presentation custom-show native IDs must be unique.");
            if (show.SlideIds.Count == 0 || show.SlideIds.Count > MaxSlidesPerShow)
                throw new CodecException(
                    "presentation_custom_show_budget_exceeded",
                    $"Presentation custom show {show.Name} must contain 1 through {MaxSlidesPerShow} slide references.");
            foreach (var slideId in show.SlideIds)
                if (!slideIds.Contains(slideId))
                    throw new CodecException(
                        "invalid_presentation_custom_show",
                        $"Presentation custom show {show.Name} references missing slide {slideId}.");

            if (!hasSourcePackage && show.Source is not null)
                throw new CodecException(
                    "invalid_presentation_custom_show",
                    $"Source-free presentation custom show {show.Name} cannot carry a source binding.");
            if (hasSourcePackage && show.Source is null)
                throw new CodecException(
                    "presentation_custom_show_topology_changed",
                    "Source-preserving PPTX export cannot add a custom show outside the imported fixed topology.",
                    "ppt/presentation.xml");
            if (show.Source is { } binding && binding.Ordinal != showIndex)
                throw new CodecException(
                    "presentation_custom_show_source_binding_mismatch",
                    $"Presentation custom show {show.Name} has an invalid source ordinal.",
                    "ppt/presentation.xml");
        }
    }

    internal static void BuildSourceFree(
        P.Presentation root,
        PresentationArtifact artifact,
        IReadOnlyDictionary<string, string> relationshipIdBySlideId)
    {
        if (artifact.CustomShows.Count == 0) return;
        var list = new P.CustomShowList();
        foreach (var show in artifact.CustomShows)
        {
            var slideList = new P.SlideList();
            foreach (var slideId in show.SlideIds)
                slideList.Append(new P.SlideListEntry { Id = relationshipIdBySlideId[slideId] });
            list.Append(new P.CustomShow(slideList) { Name = show.Name, Id = show.NativeId });
        }
        var defaultTextStyle = root.DefaultTextStyle;
        if (defaultTextStyle is null) root.Append(list);
        else root.InsertBefore(list, defaultTextStyle);
    }

    internal static bool ApplySourceBound(
        PresentationPart owner,
        PresentationArtifact requested,
        IReadOnlyDictionary<string, string> slideIdByRelationshipId,
        EffectiveCodecLimits limits)
    {
        var actual = Read(owner, slideIdByRelationshipId, limits);
        if (actual.Opaque)
        {
            if (!requested.CustomShowsOpaque || requested.CustomShows.Count != 0)
                throw new CodecException(
                    "presentation_custom_show_topology_changed",
                    "The imported custom-show graph is opaque and can only be preserved unchanged.",
                    "ppt/presentation.xml");
            return false;
        }
        if (requested.CustomShowsOpaque || actual.Shows.Count != requested.CustomShows.Count)
            throw new CodecException(
                "presentation_custom_show_topology_changed",
                $"Source-preserving PPTX export requires the original {actual.Shows.Count}-show topology.",
                "ppt/presentation.xml");
        if (actual.Shows.Count == 0) return false;

        var root = owner.Presentation!;
        var sourceElements = root.CustomShowList!.Elements<P.CustomShow>().ToArray();
        var relationshipIdBySlideId = slideIdByRelationshipId.ToDictionary(pair => pair.Value, pair => pair.Key, StringComparer.Ordinal);
        var changed = false;
        for (var showIndex = 0; showIndex < actual.Shows.Count; showIndex++)
        {
            var source = actual.Shows[showIndex];
            var target = requested.CustomShows[showIndex];
            var binding = target.Source ??
                throw new CodecException(
                    "presentation_custom_show_topology_changed",
                    $"Imported custom show {showIndex + 1} is missing its source binding.",
                    "ppt/presentation.xml");
            var sourceBinding = source.Source!;
            if (binding.Ordinal != sourceBinding.Ordinal ||
                !binding.ShowXmlSha256.Equals(sourceBinding.ShowXmlSha256, StringComparison.OrdinalIgnoreCase) ||
                !binding.SemanticSha256.Equals(sourceBinding.SemanticSha256, StringComparison.OrdinalIgnoreCase) ||
                binding.Editable != sourceBinding.Editable)
                throw new CodecException(
                    "presentation_custom_show_source_binding_mismatch",
                    $"Imported custom show {showIndex + 1} no longer matches its hash-bound source element.",
                    "ppt/presentation.xml");
            if (!target.Id.Equals(source.Id, StringComparison.Ordinal) || target.NativeId != source.NativeId)
                throw new CodecException(
                    "presentation_custom_show_topology_changed",
                    $"Imported custom show {showIndex + 1} cannot change its facade or native identity.",
                    "ppt/presentation.xml");
            if (SemanticHash(target).Equals(sourceBinding.SemanticSha256, StringComparison.OrdinalIgnoreCase)) continue;
            if (!binding.Editable)
                throw new CodecException(
                    "unsupported_presentation_custom_show_edit",
                    $"Imported custom show {showIndex + 1} is source-bound and read-only.",
                    "ppt/presentation.xml");

            var sourceElement = sourceElements[showIndex];
            sourceElement.Name = target.Name;
            var slideList = sourceElement.SlideList!;
            slideList.RemoveAllChildren<P.SlideListEntry>();
            foreach (var slideId in target.SlideIds)
                slideList.Append(new P.SlideListEntry { Id = relationshipIdBySlideId[slideId] });
            changed = true;
        }
        if (!changed) return false;

        root.Save();
        var roundTrip = Read(owner, slideIdByRelationshipId, limits);
        if (roundTrip.Opaque || roundTrip.Shows.Count != requested.CustomShows.Count)
            throw new CodecException(
                "presentation_custom_show_export_mismatch",
                "OpenChestnut could not re-read the edited custom-show graph.",
                "ppt/presentation.xml");
        for (var index = 0; index < roundTrip.Shows.Count; index++)
            if (!SemanticHash(roundTrip.Shows[index]).Equals(SemanticHash(requested.CustomShows[index]), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_custom_show_export_mismatch",
                    $"OpenChestnut custom show {index + 1} did not match the requested semantics after export.",
                    "ppt/presentation.xml");
        return true;
    }

    private static PptxCustomShowReadResult Opaque(string reason, ulong semanticItems) =>
        new([], true, semanticItems, reason);

    private static bool ValidName(string value) =>
        value.Length is > 0 and <= MaxNameLength &&
        value.Equals(value.Trim(), StringComparison.Ordinal) &&
        !value.Any(char.IsControl);

    private static bool HasExtendedAttributes(OpenXmlElement element) => element.ExtendedAttributes.Any();

    private static string SemanticHash(PresentationCustomShowArtifact show)
    {
        var semantic = show.Clone();
        semantic.Source = null;
        return Hash(semantic.ToByteArray());
    }

    private static string HashElement(OpenXmlElement element) => Hash(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
