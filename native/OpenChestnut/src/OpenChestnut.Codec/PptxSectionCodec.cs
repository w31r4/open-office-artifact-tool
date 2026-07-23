using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OpenChestnut.Codec;

internal sealed record PptxSectionReadResult(
    IReadOnlyList<PresentationSectionArtifact> Sections,
    bool Opaque,
    ulong SemanticItems,
    string? Reason);

// Owns the canonical PowerPoint 2010 p14:sectionLst extension. A section is
// deliberately stricter than a custom show: it partitions the presentation's
// ordered slide list once, so an inserted/moved slide cannot silently acquire
// a surprising section membership during a source-preserving export.
internal static class PptxSectionCodec
{
    internal const int MaxSections = 4_096;
    internal const int MaxSlidesPerSection = 16_384;
    internal const string ExtensionUri = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
    private const int MaxNameLength = 255;

    internal static PptxSectionReadResult Read(
        PresentationPart owner,
        IReadOnlyDictionary<uint, string> publicSlideIdByNativeId,
        IReadOnlyList<string> expectedSlideIds,
        EffectiveCodecLimits limits)
    {
        var root = owner.Presentation ??
            throw new CodecException("missing_presentation_root", "PPTX package has no Presentation root.", "ppt/presentation.xml");
        var extensionList = root.PresentationExtensionList;
        if (extensionList is null) return new PptxSectionReadResult([], false, 0, null);

        var extensions = extensionList.Elements<P.PresentationExtension>()
            .Where(extension => string.Equals(extension.Uri?.Value, ExtensionUri, StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (extensions.Length == 0) return new PptxSectionReadResult([], false, 0, null);
        if (extensions.Length != 1)
            return Opaque("the presentation contains multiple PowerPoint section extensions", 0);

        var extension = extensions[0];
        if (extension.ExtendedAttributes.Any() || extension.ChildElements.Count != 1 || extension.ChildElements[0] is not P14.SectionList list)
            return Opaque("the PowerPoint section extension has attributes or children outside the canonical p14:sectionLst profile", 0);
        if (list.ExtendedAttributes.Any() || list.ChildElements.Any(child => child is not P14.Section))
            return Opaque("the PowerPoint section list contains extension attributes or non-section children", checked((ulong)list.ChildElements.Count));
        if (list.ChildElements.Count is 0 or > MaxSections)
        {
            if (list.ChildElements.Count > MaxSections)
                throw new CodecException(
                    "presentation_section_budget_exceeded",
                    $"PPTX section list has {list.ChildElements.Count} children and exceeds the {MaxSections}-section budget.",
                    "ppt/presentation.xml");
            return Opaque("the PowerPoint section list is empty", 0);
        }

        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var nativeIds = new HashSet<string>(StringComparer.Ordinal);
        var sections = new List<PresentationSectionArtifact>(list.ChildElements.Count);
        ulong semanticItems = 0;
        var sourceSections = list.Elements<P14.Section>().ToArray();
        for (var sectionIndex = 0; sectionIndex < sourceSections.Length; sectionIndex++)
        {
            var source = sourceSections[sectionIndex];
            semanticItems = checked(semanticItems + 1);
            if (semanticItems > limits.MaxCells)
                throw new CodecException(
                    "presentation_item_budget_exceeded",
                    $"PPTX sections exceed max_cells semantic-item budget ({limits.MaxCells}).",
                    "ppt/presentation.xml");
            if (source.ExtendedAttributes.Any() || source.ChildElements.Count != 1 || source.ChildElements[0] is not P14.SectionSlideIdList slideList)
                return Opaque($"section {sectionIndex + 1} is not one canonical slide-ID-list child", semanticItems);

            var name = source.Name?.Value ?? string.Empty;
            var nativeId = NormalizeNativeId(source.Id?.Value);
            if (!ValidName(name) || !names.Add(name))
                return Opaque($"section {sectionIndex + 1} has a missing, invalid, or duplicate name", semanticItems);
            if (nativeId is null || !nativeIds.Add(nativeId))
                return Opaque($"section {sectionIndex + 1} has a missing, invalid, or duplicate native GUID", semanticItems);
            if (slideList.ExtendedAttributes.Any() || slideList.ChildElements.Any(child => child is not P14.SectionSlideIdListEntry))
                return Opaque($"section {sectionIndex + 1} contains extension attributes or non-slide children", semanticItems);

            var entries = slideList.Elements<P14.SectionSlideIdListEntry>().ToArray();
            if (entries.Length == 0)
                return Opaque($"section {sectionIndex + 1} has no slides", semanticItems);
            if (entries.Length > MaxSlidesPerSection)
                throw new CodecException(
                    "presentation_section_budget_exceeded",
                    $"PPTX section {sectionIndex + 1} has {entries.Length} slides and exceeds the {MaxSlidesPerSection}-slide budget.",
                    "ppt/presentation.xml");
            semanticItems = checked(semanticItems + (ulong)entries.Length);
            if (semanticItems > limits.MaxCells)
                throw new CodecException(
                    "presentation_item_budget_exceeded",
                    $"PPTX sections exceed max_cells semantic-item budget ({limits.MaxCells}).",
                    "ppt/presentation.xml");

            var section = new PresentationSectionArtifact
            {
                Id = $"section/{sectionIndex + 1}",
                Name = name,
                NativeId = nativeId,
            };
            foreach (var entry in entries)
            {
                if (entry.ExtendedAttributes.Any() || entry.Id?.Value is not { } nativeSlideId ||
                    !publicSlideIdByNativeId.TryGetValue(nativeSlideId, out var slideId))
                    return Opaque($"section {sectionIndex + 1} contains an unresolved or extended native slide ID", semanticItems);
                section.SlideIds.Add(slideId);
            }
            section.Source = new PresentationSectionSourceBinding
            {
                Ordinal = checked((uint)sectionIndex),
                SectionXmlSha256 = HashElement(source),
                Editable = true,
            };
            section.Source.SemanticSha256 = SemanticHash(section);
            sections.Add(section);
        }

        var membership = sections.SelectMany(section => section.SlideIds).ToArray();
        if (!membership.SequenceEqual(expectedSlideIds, StringComparer.Ordinal))
            return Opaque("section membership does not partition every presentation slide exactly once and in presentation order", semanticItems);
        return new PptxSectionReadResult(sections, false, semanticItems, null);
    }

    internal static void Validate(
        PresentationArtifact artifact,
        bool hasSourcePackage,
        EffectiveCodecLimits limits,
        ref ulong semanticItems)
    {
        if (artifact.SectionsOpaque)
        {
            if (!hasSourcePackage || artifact.Sections.Count != 0)
                throw new CodecException(
                    "invalid_presentation_section",
                    "Opaque PowerPoint sections require a validated source package and cannot coexist with semantic sections.");
            return;
        }
        if (artifact.Sections.Count > MaxSections)
            throw new CodecException(
                "presentation_section_budget_exceeded",
                $"Presentation cannot contain more than {MaxSections} sections.");

        var presentationSlideIds = artifact.Slides.Select(slide => slide.Id).ToArray();
        var knownSlideIds = presentationSlideIds.ToHashSet(StringComparer.Ordinal);
        if (artifact.Sections.Count > 0 && (knownSlideIds.Count != presentationSlideIds.Length || knownSlideIds.Contains(string.Empty)))
            throw new CodecException("invalid_presentation_slide", "Presentation slide IDs must be non-empty and unique when sections are present.");
        var ids = new HashSet<string>(StringComparer.Ordinal);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var nativeIds = new HashSet<string>(StringComparer.Ordinal);
        var membership = new List<string>();
        for (var sectionIndex = 0; sectionIndex < artifact.Sections.Count; sectionIndex++)
        {
            var section = artifact.Sections[sectionIndex];
            semanticItems = checked(semanticItems + 1 + (ulong)section.SlideIds.Count);
            if (semanticItems > limits.MaxCells)
                throw new CodecException(
                    "presentation_item_budget_exceeded",
                    $"Presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).");
            if (string.IsNullOrWhiteSpace(section.Id) || section.Id.Length > 1_024 || !ids.Add(section.Id))
                throw new CodecException("invalid_presentation_section", "Presentation section IDs must be non-empty, bounded, and unique.");
            if (!ValidName(section.Name) || !names.Add(section.Name))
                throw new CodecException("invalid_presentation_section", "Presentation section names must be trimmed, valid, and case-insensitively unique.");
            var nativeId = NormalizeNativeId(section.NativeId);
            if (nativeId is null || !nativeId.Equals(section.NativeId, StringComparison.Ordinal) || !nativeIds.Add(nativeId))
                throw new CodecException("invalid_presentation_section", "Presentation section native IDs must be unique canonical brace-delimited GUIDs.");
            if (section.SlideIds.Count is 0 or > MaxSlidesPerSection)
                throw new CodecException(
                    "presentation_section_budget_exceeded",
                    $"Presentation section {section.Name} must contain 1 through {MaxSlidesPerSection} slide references.");
            foreach (var slideId in section.SlideIds)
            {
                if (!knownSlideIds.Contains(slideId))
                    throw new CodecException(
                        "invalid_presentation_section",
                        $"Presentation section {section.Name} references missing slide {slideId}.");
                membership.Add(slideId);
            }

            if (!hasSourcePackage && section.Source is not null)
                throw new CodecException(
                    "invalid_presentation_section",
                    $"Source-free presentation section {section.Name} cannot carry a source binding.");
            if (hasSourcePackage && section.Source is null)
                throw new CodecException(
                    "presentation_section_topology_changed",
                    "Source-preserving PPTX export cannot add a PowerPoint section outside the imported fixed topology.",
                    "ppt/presentation.xml");
            if (section.Source is { } binding && binding.Ordinal != sectionIndex)
                throw new CodecException(
                    "presentation_section_source_binding_mismatch",
                    $"Presentation section {section.Name} has an invalid source ordinal.",
                    "ppt/presentation.xml");
        }
        // Pending source-bound clones are rejected later by the dedicated
        // clone preflight. Deferring only this aggregate check lets that
        // operation report its true incompatibility instead of pretending the
        // caller authored an ordinary malformed section partition.
        var hasPendingClone = artifact.Slides.Any(slide => slide.CloneSource is not null);
        if (artifact.Sections.Count > 0 && !hasPendingClone && !membership.SequenceEqual(presentationSlideIds, StringComparer.Ordinal))
            throw new CodecException(
                "invalid_presentation_section",
                "Presentation sections must partition every slide exactly once and in presentation order.",
                "ppt/presentation.xml");
    }

    internal static void BuildSourceFree(
        P.Presentation root,
        PresentationArtifact artifact,
        IReadOnlyDictionary<string, uint> nativeSlideIdByPublicId)
    {
        if (artifact.Sections.Count == 0) return;
        var list = new P14.SectionList();
        foreach (var section in artifact.Sections)
        {
            var slideList = new P14.SectionSlideIdList();
            foreach (var slideId in section.SlideIds)
                slideList.Append(new P14.SectionSlideIdListEntry { Id = nativeSlideIdByPublicId[slideId] });
            list.Append(new P14.Section(slideList) { Name = section.Name, Id = section.NativeId });
        }
        var extension = new P.PresentationExtension(list) { Uri = ExtensionUri };
        var extensionList = root.PresentationExtensionList;
        if (extensionList is null) root.Append(new P.PresentationExtensionList(extension));
        else extensionList.Append(extension);
    }

    internal static bool ApplySourceBound(
        PresentationPart owner,
        PresentationArtifact requested,
        IReadOnlyDictionary<uint, string> publicSlideIdByNativeId,
        IReadOnlyList<string> expectedSlideIds,
        EffectiveCodecLimits limits)
    {
        var actual = Read(owner, publicSlideIdByNativeId, expectedSlideIds, limits);
        if (actual.Opaque)
        {
            if (!requested.SectionsOpaque || requested.Sections.Count != 0)
                throw new CodecException(
                    "presentation_section_topology_changed",
                    "The imported PowerPoint section graph is opaque and can only be preserved unchanged.",
                    "ppt/presentation.xml");
            return false;
        }
        if (requested.SectionsOpaque || actual.Sections.Count != requested.Sections.Count)
            throw new CodecException(
                "presentation_section_topology_changed",
                $"Source-preserving PPTX export requires the original {actual.Sections.Count}-section topology.",
                "ppt/presentation.xml");
        if (actual.Sections.Count == 0) return false;

        var root = owner.Presentation!;
        var extension = SectionExtension(root) ?? throw new CodecException(
            "presentation_section_topology_changed",
            "The imported PowerPoint section extension disappeared before export.",
            "ppt/presentation.xml");
        var sourceElements = extension.GetFirstChild<P14.SectionList>()!.Elements<P14.Section>().ToArray();
        var nativeSlideIdByPublicId = publicSlideIdByNativeId
            .ToDictionary(pair => pair.Value, pair => pair.Key, StringComparer.Ordinal);
        var changed = false;
        for (var sectionIndex = 0; sectionIndex < actual.Sections.Count; sectionIndex++)
        {
            var source = actual.Sections[sectionIndex];
            var target = requested.Sections[sectionIndex];
            var binding = target.Source ??
                throw new CodecException(
                    "presentation_section_topology_changed",
                    $"Imported PowerPoint section {sectionIndex + 1} is missing its source binding.",
                    "ppt/presentation.xml");
            var sourceBinding = source.Source!;
            if (binding.Ordinal != sourceBinding.Ordinal ||
                !binding.SectionXmlSha256.Equals(sourceBinding.SectionXmlSha256, StringComparison.OrdinalIgnoreCase) ||
                !binding.SemanticSha256.Equals(sourceBinding.SemanticSha256, StringComparison.OrdinalIgnoreCase) ||
                binding.Editable != sourceBinding.Editable)
                throw new CodecException(
                    "presentation_section_source_binding_mismatch",
                    $"Imported PowerPoint section {sectionIndex + 1} no longer matches its hash-bound source element.",
                    "ppt/presentation.xml");
            if (!target.Id.Equals(source.Id, StringComparison.Ordinal) || !target.NativeId.Equals(source.NativeId, StringComparison.Ordinal))
                throw new CodecException(
                    "presentation_section_topology_changed",
                    $"Imported PowerPoint section {sectionIndex + 1} cannot change its facade or native GUID identity.",
                    "ppt/presentation.xml");
            if (SemanticHash(target).Equals(sourceBinding.SemanticSha256, StringComparison.OrdinalIgnoreCase)) continue;
            if (!binding.Editable)
                throw new CodecException(
                    "unsupported_presentation_section_edit",
                    $"Imported PowerPoint section {sectionIndex + 1} is source-bound and read-only.",
                    "ppt/presentation.xml");

            var sourceElement = sourceElements[sectionIndex];
            sourceElement.Name = target.Name;
            var slideList = sourceElement.SectionSlideIdList!;
            slideList.RemoveAllChildren<P14.SectionSlideIdListEntry>();
            foreach (var slideId in target.SlideIds)
                slideList.Append(new P14.SectionSlideIdListEntry { Id = nativeSlideIdByPublicId[slideId] });
            changed = true;
        }
        if (!changed) return false;

        root.Save();
        var roundTrip = Read(owner, publicSlideIdByNativeId, expectedSlideIds, limits);
        if (roundTrip.Opaque || roundTrip.Sections.Count != requested.Sections.Count)
            throw new CodecException(
                "presentation_section_export_mismatch",
                "OpenChestnut could not re-read the edited PowerPoint section graph.",
                "ppt/presentation.xml");
        for (var index = 0; index < roundTrip.Sections.Count; index++)
            if (!SemanticHash(roundTrip.Sections[index]).Equals(SemanticHash(requested.Sections[index]), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_section_export_mismatch",
                    $"OpenChestnut section {index + 1} did not match the requested semantics after export.",
                    "ppt/presentation.xml");
        return true;
    }

    internal static void AssertNoSectionCloneCombination(
        PresentationPart owner,
        PresentationArtifact requested,
        IReadOnlyDictionary<uint, string> publicSlideIdByNativeId,
        IReadOnlyList<string> expectedSlideIds,
        EffectiveCodecLimits limits)
    {
        var actual = Read(owner, publicSlideIdByNativeId, expectedSlideIds, limits);
        if (actual.Opaque || actual.Sections.Count > 0 || requested.SectionsOpaque || requested.Sections.Count > 0)
            throw new CodecException(
                "unsupported_presentation_slide_clone",
                "The bounded slide-clone profile cannot combine a clone with PowerPoint sections; export and reimport a deck without sections or use an explicit OPC graph operation.",
                "ppt/presentation.xml");
    }

    private static P.PresentationExtension? SectionExtension(P.Presentation root) => root.PresentationExtensionList?
        .Elements<P.PresentationExtension>()
        .SingleOrDefault(extension => string.Equals(extension.Uri?.Value, ExtensionUri, StringComparison.OrdinalIgnoreCase));

    private static PptxSectionReadResult Opaque(string reason, ulong semanticItems) =>
        new([], true, semanticItems, reason);

    private static bool ValidName(string value) =>
        value.Length is > 0 and <= MaxNameLength &&
        value.Equals(value.Trim(), StringComparison.Ordinal) &&
        !value.Any(char.IsControl);

    private static string? NormalizeNativeId(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || !Guid.TryParseExact(value, "B", out var guid)) return null;
        return guid.ToString("B").ToUpperInvariant();
    }

    private static string SemanticHash(PresentationSectionArtifact section)
    {
        var semantic = section.Clone();
        semantic.Source = null;
        return Hash(semantic.ToByteArray());
    }

    private static string HashElement(OpenXmlElement element) => Hash(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
