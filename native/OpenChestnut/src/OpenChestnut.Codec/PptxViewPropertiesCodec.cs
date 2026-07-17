using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenChestnut.Codec;

// Imports presentation-wide grid/snap/guide metadata as a read-only semantic
// view. JavaScript editor visibility is intentionally not written to the file.
internal static class PptxViewPropertiesCodec
{
    private const int MaxGuides = 1_024;

    internal static PresentationViewProperties? Read(PresentationPart owner)
    {
        var part = owner.ViewPropertiesPart;
        if (part is null) return null;
        var root = part.ViewProperties ??
            throw new CodecException("missing_presentation_view_root", "PPTX view-properties part has no p:viewPr root.", PartPath(part));
        var result = new PresentationViewProperties();
        if (root.GridSpacing?.Cx?.Value is { } cx) result.GridSpacingCxEmu = cx;
        if (root.GridSpacing?.Cy?.Value is { } cy) result.GridSpacingCyEmu = cy;
        var common = root.SlideViewProperties?.CommonSlideViewProperties;
        if (common?.SnapToGrid?.Value is { } snapToGrid) result.SlideViewSnapToGrid = snapToGrid;
        if (common?.SnapToObjects?.Value is { } snapToObjects) result.SlideViewSnapToObjects = snapToObjects;
        foreach (var guide in common?.GuideList?.Elements<P.Guide>() ?? [])
        {
            if (result.SlideGuides.Count >= MaxGuides)
                throw new CodecException(
                    "presentation_guide_budget_exceeded",
                    $"PPTX presentation view exceeds the {MaxGuides}-guide budget.",
                    PartPath(part));
            if (guide.Position?.Value is not { } position) continue;
            result.SlideGuides.Add(new PresentationSlideGuide
            {
                Orientation = guide.Orientation?.Value == P.DirectionValues.Vertical
                    ? PresentationSlideGuide.Types.Orientation.Vertical
                    : PresentationSlideGuide.Types.Orientation.Horizontal,
                Position = position,
            });
        }
        result.Source = new PresentationViewPropertiesSourceBinding
        {
            PartPath = PartPath(part),
            RelationshipId = owner.GetIdOfPart(part),
            ViewXmlSha256 = HashElement(root),
            SemanticSha256 = SemanticHash(result),
        };
        return result;
    }

    internal static void AssertSource(PresentationPart owner, PresentationViewProperties? requested)
    {
        var actual = Read(owner);
        if (actual is null && requested is null) return;
        if (actual is null || requested?.Source is null)
            throw new CodecException(
                "presentation_view_topology_changed",
                "Source-preserving PPTX export requires the original presentation view-properties topology.",
                "ppt/presentation.xml");
        var binding = requested.Source;
        if (!binding.PartPath.Equals(actual.Source.PartPath, StringComparison.OrdinalIgnoreCase) ||
            !binding.RelationshipId.Equals(actual.Source.RelationshipId, StringComparison.Ordinal) ||
            !binding.ViewXmlSha256.Equals(actual.Source.ViewXmlSha256, StringComparison.OrdinalIgnoreCase) ||
            !binding.SemanticSha256.Equals(actual.Source.SemanticSha256, StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "presentation_view_source_binding_mismatch",
                "Presentation view properties no longer match their source integrity binding.",
                actual.Source.PartPath);
        if (!SemanticHash(requested).Equals(SemanticHash(actual), StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "unsupported_presentation_view_edit",
                "Imported presentation grid spacing, snap settings, and guides are source-bound and read-only.",
                actual.Source.PartPath);
    }

    internal static void Validate(PresentationViewProperties? properties, bool hasSourcePackage)
    {
        if (properties is null) return;
        if (!hasSourcePackage || properties.Source is null)
            throw new CodecException(
                "unsupported_presentation_features",
                "Source-free authoring of PowerPoint view properties and guides is unsupported; use presentation.view for local editor visibility.");
        if ((properties.HasGridSpacingCxEmu && properties.GridSpacingCxEmu <= 0) ||
            (properties.HasGridSpacingCyEmu && properties.GridSpacingCyEmu <= 0))
            throw new CodecException("invalid_presentation_view", "Presentation grid spacing must be positive when present.");
        foreach (var guide in properties.SlideGuides)
        {
            if (guide.Orientation is not PresentationSlideGuide.Types.Orientation.Horizontal and
                not PresentationSlideGuide.Types.Orientation.Vertical)
                throw new CodecException("invalid_presentation_view", "Presentation guides require horizontal or vertical orientation.");
        if (properties.SlideGuides.Count > MaxGuides)
            throw new CodecException("presentation_guide_budget_exceeded", $"Presentation cannot contain more than {MaxGuides} guides.");
        }
    }

    private static string SemanticHash(PresentationViewProperties properties)
    {
        var semantic = properties.Clone();
        semantic.Source = null;
        return Hash(semantic.ToByteArray());
    }

    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static string HashElement(OpenXmlElement element) => Hash(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
