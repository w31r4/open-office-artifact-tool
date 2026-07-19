using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenChestnut.Codec;

internal sealed record PptxImportResult(ArtifactEnvelope Artifact, IReadOnlyList<Diagnostic> Diagnostics);
internal sealed record PptxExportResult(byte[] File, IReadOnlyList<Diagnostic> Diagnostics);
internal sealed record PptxLayoutGraphEntry(int Index, string Id, string RelationshipId, SlideLayoutPart Part);
internal sealed record PptxMasterGraphEntry(int Index, string Id, string RelationshipId, SlideMasterPart Part, IReadOnlyList<PptxLayoutGraphEntry> Layouts);
internal sealed record PptxSourceSlideEntry(int Index, P.SlideId SlideId, string RelationshipId, SlidePart Part);
// Clone preflight produces this closed graph plan once. The execution path
// consumes only this plan so adding a new supported leaf cannot create a
// second, subtly different relationship inventory after validation.
internal sealed record PptxCloneImageEdge(ImagePart Part, string RelationshipId);
internal sealed record PptxCloneExternalHyperlinkEdge(Uri Uri, string RelationshipId);
internal sealed record PptxCloneSlideHyperlinkEdge(SlidePart Part, string RelationshipId);
internal sealed record PptxCloneNotesLeaf(
    NotesSlidePart Part,
    string SlideRelationshipId,
    NotesMasterPart MasterPart,
    string MasterRelationshipId,
    string SlideBackReferenceRelationshipId);
// Legacy comments are a clone-local XML leaf whose author IDs resolve through
// the presentation-wide immutable catalog. The catalog has no direct OPC edge
// from the comment part, so keep it explicit in the preflight plan rather than
// treating that semantic dependency as an incidental implementation detail.
internal sealed record PptxCloneLegacyCommentsLeaf(
    SlideCommentsPart Part,
    string SlideRelationshipId,
    CommentAuthorsPart AuthorsPart,
    string AuthorsRelationshipId);
internal sealed record PptxCloneLeaf(
    SlideLayoutPart LayoutPart,
    string LayoutRelationshipId,
    IReadOnlyList<PptxCloneImageEdge> Images,
    IReadOnlyList<PptxCloneExternalHyperlinkEdge> ExternalHyperlinks,
    IReadOnlyList<PptxCloneSlideHyperlinkEdge> SlideHyperlinks,
    PptxCloneNotesLeaf? Notes,
    PptxCloneLegacyCommentsLeaf? LegacyComments);
internal sealed class PptxTargetSlideEntry
{
    internal PptxTargetSlideEntry(int targetIndex, PresentationSlide target, PptxSourceSlideEntry source, bool isClone)
    {
        TargetIndex = targetIndex;
        Target = target;
        Source = source;
        IsClone = isClone;
        OutputSlideId = source.SlideId;
        OutputPart = source.Part;
    }

    internal int TargetIndex { get; }
    internal PresentationSlide Target { get; }
    internal PptxSourceSlideEntry Source { get; }
    internal bool IsClone { get; }
    internal P.SlideId OutputSlideId { get; set; }
    internal SlidePart OutputPart { get; set; }
}

internal static class PptxCodec
{
    private const long DefaultSlideWidthEmu = 12_192_000;
    private const long DefaultSlideHeightEmu = 6_858_000;
    private const int MaxCloneHyperlinkRelationships = 4_096;

    internal static PptxImportResult Import(byte[] bytes, EffectiveCodecLimits limits)
    {
        var opaque = PackageGuards.ValidateAndCollectOpaque(bytes, limits, OpcPackageProfile.Pptx);
        var nativeObjects = new PptxNativeObjectCatalog(opaque, bytes, limits);
        var diagnostics = new List<Diagnostic>();
        var opaqueCount = opaque.Parts.Count + opaque.PackageRelationships.Count;
        if (opaqueCount > 0)
            diagnostics.Add(CodecProtocol.Warning(
                "opaque_content_retained",
                $"Retained {opaqueCount} unsupported OPC parts or relationships for source-bound, fail-closed export from the validated package snapshot.",
                opaque.Parts.FirstOrDefault()?.Path ?? opaque.PackageRelationships.FirstOrDefault()?.SourcePath));

        using var stream = new MemoryStream(bytes, writable: false);
        using var package = PresentationDocument.Open(stream, isEditable: false);
        var presentationPart = package.PresentationPart ??
            throw new CodecException("missing_presentation_part", "PPTX package has no Presentation part.", "ppt/presentation.xml");
        var presentationRoot = presentationPart.Presentation ??
            throw new CodecException("missing_presentation_root", "PPTX package has no presentation root.", "ppt/presentation.xml");
        var slideIds = presentationRoot.SlideIdList?.Elements<P.SlideId>().ToArray() ?? [];
        if ((uint)slideIds.Length > limits.MaxSheets)
            throw new CodecException("slide_budget_exceeded", $"PPTX presentation has {slideIds.Length} slides and exceeds max_sheets ({limits.MaxSheets}).", "ppt/presentation.xml");
        var slideParts = ResolveSlideParts(presentationPart, slideIds);
        var slideIdByPartPath = slideParts
            .Select((part, index) => (Path: PartPath(part), Id: $"presentation/slide/{index + 1}"))
            .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
        var assetCatalog = new PptxAssetCatalog([], limits);
        var masterGraph = ReadMasterGraph(presentationPart);
        var layoutIdByPartPath = masterGraph
            .SelectMany(master => master.Layouts)
            .ToDictionary(layout => PartPath(layout.Part), layout => layout.Id, StringComparer.OrdinalIgnoreCase);

        var artifact = new PresentationArtifact
        {
            Id = "presentation/1",
            Name = "Imported presentation",
            SlideWidthEmu = presentationRoot.SlideSize?.Cx?.Value ?? DefaultSlideWidthEmu,
            SlideHeightEmu = presentationRoot.SlideSize?.Cy?.Value ?? DefaultSlideHeightEmu,
        };
        artifact.ViewProperties = PptxViewPropertiesCodec.Read(presentationPart);
        foreach (var master in masterGraph)
        {
            var masterRoot = master.Part.SlideMaster ??
                throw new CodecException("missing_slide_master_root", $"Presentation master {master.Index + 1} has no slide master root.", PartPath(master.Part));
            var masterCommon = masterRoot.CommonSlideData ??
                throw new CodecException("missing_common_slide_data", $"Presentation master {master.Index + 1} has no common slide data.", PartPath(master.Part));
            var masterShapeTree = masterCommon.ShapeTree ??
                throw new CodecException("missing_shape_tree", $"Presentation master {master.Index + 1} has no shape tree.", PartPath(master.Part));
            var masterContext = new PptxPartContext(master.Part, slideIdByPartPath, assets: assetCatalog);
            var textStyles = PptxMasterTextStylesCodec.Read(masterRoot, masterContext);
            var background = PptxBackgroundCodec.Read(masterCommon);
            var masterArtifact = new PresentationMaster
            {
                Id = master.Id,
                Name = masterCommon.Name?.Value ?? $"Master {master.Index + 1}",
                TextStyles = textStyles,
                Source = new PresentationMasterSourceBinding
                {
                    MasterIndex = checked((uint)master.Index),
                    PartPath = PartPath(master.Part),
                    RelationshipId = master.RelationshipId,
                    MasterXmlSha256 = HashElement(masterRoot),
                    TextStylesSemanticSha256 = MasterTextStylesSemanticHash(textStyles),
                    TextStylesEditable = PptxMasterTextStylesCodec.Supports(masterRoot),
                    BackgroundSemanticSha256 = BackgroundSemanticHash(background),
                    BackgroundEditable = PptxBackgroundCodec.Supports(masterCommon),
                },
            };
            if (background is not null) masterArtifact.Background = background;
            masterArtifact.Placeholders.Add(PptxPlaceholderCodec.Read(masterShapeTree, master.Id, masterContext));
            artifact.Masters.Add(masterArtifact);
            foreach (var layout in master.Layouts)
            {
                var layoutRoot = layout.Part.SlideLayout ??
                    throw new CodecException("missing_slide_layout_root", $"Presentation layout {layout.Index + 1} under master {master.Index + 1} has no slide layout root.", PartPath(layout.Part));
                var layoutCommon = layoutRoot.CommonSlideData ??
                    throw new CodecException("missing_common_slide_data", $"Presentation layout {layout.Index + 1} under master {master.Index + 1} has no common slide data.", PartPath(layout.Part));
                var layoutShapeTree = layoutCommon.ShapeTree ??
                    throw new CodecException("missing_shape_tree", $"Presentation layout {layout.Index + 1} under master {master.Index + 1} has no shape tree.", PartPath(layout.Part));
                var layoutContext = new PptxPartContext(layout.Part, slideIdByPartPath, assets: assetCatalog);
                var layoutBackground = PptxBackgroundCodec.Read(layoutCommon);
                var layoutArtifact = new PresentationLayout
                {
                    Id = layout.Id,
                    Name = layoutCommon.Name?.Value ?? $"Layout {layout.Index + 1}",
                    MasterId = master.Id,
                    Type = LayoutTypeName(layoutRoot),
                    Source = new PresentationLayoutSourceBinding
                    {
                        LayoutIndex = checked((uint)layout.Index),
                        PartPath = PartPath(layout.Part),
                        RelationshipId = layout.RelationshipId,
                        LayoutXmlSha256 = HashElement(layoutRoot),
                        BackgroundSemanticSha256 = BackgroundSemanticHash(layoutBackground),
                        BackgroundEditable = PptxBackgroundCodec.Supports(layoutCommon),
                    },
                };
                if (layoutBackground is not null) layoutArtifact.Background = layoutBackground;
                layoutArtifact.Placeholders.Add(PptxPlaceholderCodec.Read(layoutShapeTree, layout.Id, layoutContext));
                artifact.Layouts.Add(layoutArtifact);
            }
        }
        ulong semanticItems = 0;
        for (var slideIndex = 0; slideIndex < slideIds.Length; slideIndex++)
        {
            var slideId = slideIds[slideIndex];
            var relationshipId = slideId.RelationshipId?.Value ?? string.Empty;
            var slidePart = slideParts[slideIndex];
            var slideRoot = slidePart.Slide ??
                throw new CodecException("missing_slide_root", $"Presentation slide {slideIndex + 1} has no slide root.", PartPath(slidePart));
            var slideCommon = slideRoot.CommonSlideData ??
                throw new CodecException("missing_common_slide_data", $"Presentation slide {slideIndex + 1} has no common slide data.", PartPath(slidePart));
            var shapeTree = slideCommon.ShapeTree ??
                throw new CodecException("missing_shape_tree", $"Presentation slide {slideIndex + 1} has no shape tree.", PartPath(slidePart));
            var slideBackground = PptxBackgroundCodec.Read(slideCommon);
            var elements = ShapeElements(shapeTree);
            var slideArtifactId = $"presentation/slide/{slideIndex + 1}";
            var elementIdsByNativeId = NativeElementIds(elements, slideArtifactId);
            var target = new PresentationSlide
            {
                Id = slideArtifactId,
                Name = slideRoot.CommonSlideData?.Name?.Value ?? $"Slide {slideIndex + 1}",
                LayoutId = slidePart.SlideLayoutPart is { } layoutPart
                    ? layoutIdByPartPath.GetValueOrDefault(PartPath(layoutPart)) ??
                      throw new CodecException("unresolved_slide_layout_binding", $"Presentation slide {slideIndex + 1} references a layout outside the master graph.", PartPath(slidePart))
                    : string.Empty,
                Source = new PresentationSlideSourceBinding
                {
                    SlideIndex = checked((uint)slideIndex),
                    PartPath = PartPath(slidePart),
                    RelationshipId = relationshipId,
                    SlideXmlSha256 = HashElement(slideRoot),
                    LayoutRelationshipId = slidePart.SlideLayoutPart is { } boundLayout ? slidePart.GetIdOfPart(boundLayout) : string.Empty,
                    BackgroundSemanticSha256 = BackgroundSemanticHash(slideBackground),
                    BackgroundEditable = PptxBackgroundCodec.Supports(slideCommon),
                },
            };
            if (slideBackground is not null) target.Background = slideBackground;
            if (PptxSpeakerNotesCodec.Read(slidePart) is { } speakerNotes)
                target.SpeakerNotes = speakerNotes;
            target.LegacyComments.Add(PptxLegacyCommentsCodec.Read(presentationPart, slidePart, slideIndex, diagnostics));
            target.ModernComments.Add(PptxModernCommentsCodec.Read(
                presentationPart,
                slideId,
                slidePart,
                elements,
                elementIdsByNativeId,
                slideIndex,
                diagnostics));
            var slideContext = new PptxPartContext(slidePart, slideIdByPartPath, assets: assetCatalog);
            for (var elementIndex = 0; elementIndex < elements.Length; elementIndex++)
            {
                semanticItems++;
                if (semanticItems > limits.MaxCells)
                    throw new CodecException("presentation_item_budget_exceeded", $"PPTX presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).", PartPath(slidePart));
                var importedElement = ReadElement(elements[elementIndex], slideIndex, elementIndex, slideContext, nativeObjects, elementIdsByNativeId);
                if (importedElement.ContentCase == PresentationElement.ContentOneofCase.Table)
                {
                    semanticItems += checked((ulong)importedElement.Table.Rows.Sum(row => row.Cells.Count));
                    if (semanticItems > limits.MaxCells)
                        throw new CodecException("presentation_item_budget_exceeded", $"PPTX presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).", PartPath(slidePart));
                }
                target.Elements.Add(importedElement);
            }
            artifact.Slides.Add(target);
        }

        var envelope = new ArtifactEnvelope
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Family = ArtifactFamily.Presentation,
            Presentation = artifact,
            OpaqueOpc = opaque,
            Source = new SourceIdentity
            {
                Format = "pptx",
                PackageSha256 = Hash(bytes),
                Producer = "open-office-artifact-tool/OpenChestnut",
            },
        };
        envelope.Assets.Add(assetCatalog.ImportedAssets);
        envelope.Diagnostics.Add(diagnostics);
        return new PptxImportResult(envelope, diagnostics);
    }

    internal static PptxExportResult Export(ArtifactEnvelope envelope, EffectiveCodecLimits limits)
    {
        var requiresSourcePreservation =
            envelope.ProtocolVersion == CodecProtocol.ProtocolVersion &&
            envelope.Family == ArtifactFamily.Presentation &&
            envelope.PayloadCase == ArtifactEnvelope.PayloadOneofCase.Presentation &&
            RequiresSourcePreservation(envelope);
        if (requiresSourcePreservation && envelope.OpaqueOpc?.SourcePackage is not { Data.IsEmpty: false })
            throw new CodecException(
                "missing_source_package",
                "Source-bound PPTX export requires its validated source package snapshot.");

        var assetCatalog = ValidateEnvelope(envelope, limits);
        var opaqueCount = (envelope.OpaqueOpc?.Parts.Count ?? 0) +
                          (envelope.OpaqueOpc?.PackageRelationships.Count ?? 0);
        if (requiresSourcePreservation)
            return ExportPreservingSource(envelope, limits, opaqueCount, assetCatalog);

        var diagnostics = new List<Diagnostic>();

        using var stream = new MemoryStream();
        using (var package = PresentationDocument.Create(stream, PresentationDocumentType.Presentation, autoSave: true))
            BuildPresentation(package, envelope.Presentation, assetCatalog);
        var bytes = stream.ToArray();
        ValidateOutputBudget(bytes, limits);
        ValidateOffice2021(bytes);
        return new PptxExportResult(bytes, diagnostics);
    }

    private static bool RequiresSourcePreservation(ArtifactEnvelope envelope)
    {
        if (envelope.Source is not null) return true;
        if (envelope.OpaqueOpc is { } opaque &&
            (opaque.SourcePackage is not null || opaque.Parts.Count > 0 || opaque.PackageRelationships.Count > 0))
            return true;

        var presentation = envelope.Presentation;
        return presentation.Masters.Any(master =>
                   master.Source is not null || master.Placeholders.Any(placeholder => placeholder.Source is not null)) ||
               presentation.Layouts.Any(layout =>
                   layout.Source is not null || layout.Placeholders.Any(placeholder => placeholder.Source is not null)) ||
               presentation.Slides.Any(slide =>
                   slide.Source is not null || slide.Elements.Any(element =>
                       element.Source is not null || element.ContentCase == PresentationElement.ContentOneofCase.Opaque)) ||
               presentation.ViewProperties?.Source is not null;
    }

    private static PptxExportResult ExportPreservingSource(ArtifactEnvelope envelope, EffectiveCodecLimits limits, int opaqueCount, PptxAssetCatalog assetCatalog)
    {
        var sourceBytes = PackageGuards.ValidateSourcePackage(envelope.OpaqueOpc, envelope.Source, limits, OpcPackageProfile.Pptx);
        var nativeObjects = new PptxNativeObjectCatalog(envelope.OpaqueOpc, sourceBytes, limits);
        using var stream = new MemoryStream();
        stream.Write(sourceBytes);
        stream.Position = 0;
        var changedParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var addedRelationshipIds = new HashSet<string>(StringComparer.Ordinal);
        var addedPartPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var replacedOpaquePartHashes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var removedSourceSlidePartPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using (var package = PresentationDocument.Open(stream, isEditable: true, new OpenSettings { AutoSave = false }))
        {
            var presentationPart = package.PresentationPart ??
                throw new CodecException("missing_presentation_part", "PPTX package has no Presentation part.", "ppt/presentation.xml");
            var slideIds = presentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>().ToArray() ?? [];
            var targetSlides = BindSourcePreservingSlides(presentationPart, slideIds, envelope.Presentation.Slides);
            AssertCloneOriginsRetained(targetSlides);
            var retainedTargets = targetSlides.Where(target => !target.IsClone).ToArray();
            var slideParts = retainedTargets.Select(target => target.Source.Part).ToArray();
            var slideIdByPartPath = retainedTargets
                .Select(target => (Path: PartPath(target.Source.Part), Id: target.Target.Id))
                .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
            var slidePartById = retainedTargets
                .Select(target => (Part: target.Source.Part, Id: target.Target.Id))
                .ToDictionary(item => item.Id, item => item.Part, StringComparer.Ordinal);
            var masterGraph = ReadMasterGraph(presentationPart);
            if (masterGraph.Length != envelope.Presentation.Masters.Count)
                throw new CodecException(
                    "presentation_master_topology_changed",
                    $"Source-preserving PPTX export requires the original {masterGraph.Length}-master topology; the artifact contains {envelope.Presentation.Masters.Count} masters.",
                    "ppt/presentation.xml");
            var layoutGraph = masterGraph.SelectMany(master => master.Layouts.Select(layout => (Master: master, Layout: layout))).ToArray();
            if (layoutGraph.Length != envelope.Presentation.Layouts.Count)
                throw new CodecException(
                    "presentation_layout_topology_changed",
                    $"Source-preserving PPTX export requires the original {layoutGraph.Length}-layout topology; the artifact contains {envelope.Presentation.Layouts.Count} layouts.",
                    "ppt/presentation.xml");
            var layoutIdByPartPath = layoutGraph.ToDictionary(item => PartPath(item.Layout.Part), item => item.Layout.Id, StringComparer.OrdinalIgnoreCase);
            PptxViewPropertiesCodec.AssertSource(presentationPart, envelope.Presentation.ViewProperties);
            PptxLegacyCommentsCodec.AssertSourceUnchanged(presentationPart, slideParts, retainedTargets.Select(target => target.Target).ToArray());
            CloneRequestedSourceSlides(
                presentationPart,
                targetSlides,
                layoutIdByPartPath,
                slideIdByPartPath,
                assetCatalog,
                changedParts,
                addedRelationshipIds,
                addedPartPaths);
            DeleteUnrequestedSourceSlides(
                presentationPart,
                slideIds,
                targetSlides,
                envelope.OpaqueOpc,
                changedParts,
                removedSourceSlidePartPaths);
            if (ReorderSourceSlideIdList(presentationPart, targetSlides))
                changedParts.Add(PartPath(presentationPart));
            assetCatalog.IndexExistingParts(slideParts.SelectMany(part => part.ImageParts)
                .Concat(masterGraph.SelectMany(master => master.Part.Parts.Select(pair => pair.OpenXmlPart).OfType<ImagePart>())));

            for (var masterIndex = 0; masterIndex < masterGraph.Length; masterIndex++)
            {
                var graph = masterGraph[masterIndex];
                var masterRoot = graph.Part.SlideMaster ??
                    throw new CodecException("missing_slide_master_root", $"Presentation master {masterIndex + 1} has no slide master root.", PartPath(graph.Part));
                var masterCommon = masterRoot.CommonSlideData ??
                    throw new CodecException("missing_common_slide_data", $"Presentation master {masterIndex + 1} has no common slide data.", PartPath(graph.Part));
                var masterShapeTree = masterCommon.ShapeTree ??
                    throw new CodecException("missing_shape_tree", $"Presentation master {masterIndex + 1} has no shape tree.", PartPath(graph.Part));
                var target = envelope.Presentation.Masters[masterIndex];
                var binding = target.Source ?? throw new CodecException(
                    "missing_presentation_master_binding",
                    $"Presentation master {masterIndex + 1} is missing its source binding.",
                    "ppt/presentation.xml");
                if (target.Id != graph.Id ||
                    binding.MasterIndex != masterIndex ||
                    !binding.PartPath.Equals(PartPath(graph.Part), StringComparison.OrdinalIgnoreCase) ||
                    !binding.RelationshipId.Equals(graph.RelationshipId, StringComparison.Ordinal) ||
                    !binding.MasterXmlSha256.Equals(HashElement(masterRoot), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_master_binding_mismatch",
                        $"Presentation master {masterIndex + 1} does not match its hash-bound source master.",
                        PartPath(graph.Part));
                var sourceName = masterCommon.Name?.Value ?? $"Master {masterIndex + 1}";
                if (!target.Name.Equals(sourceName, StringComparison.Ordinal))
                    throw new CodecException("unsupported_presentation_edit", $"Source-preserving PPTX export cannot rename master {masterIndex + 1}.", PartPath(graph.Part));
                var masterContext = new PptxPartContext(graph.Part, slideIdByPartPath, slidePartById, assetCatalog);
                var originalStyles = PptxMasterTextStylesCodec.Read(masterRoot, masterContext);
                var originalSemanticHash = MasterTextStylesSemanticHash(originalStyles);
                if (!binding.TextStylesSemanticSha256.Equals(originalSemanticHash, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_master_source_semantics_mismatch",
                        $"Presentation master {masterIndex + 1} text styles do not match their source binding.",
                        PartPath(graph.Part));
                var requestedSemanticHash = MasterTextStylesSemanticHash(target.TextStyles);
                if (!requestedSemanticHash.Equals(originalSemanticHash, StringComparison.OrdinalIgnoreCase))
                {
                    if (!binding.TextStylesEditable || !PptxMasterTextStylesCodec.Supports(masterRoot))
                        throw new CodecException("unsupported_presentation_edit", $"Presentation master {masterIndex + 1} text styles are preserved but not safely editable by this codec slice.", PartPath(graph.Part));
                    PptxMasterTextStylesCodec.Apply(masterRoot, target.TextStyles ?? new PresentationMasterTextStyles(), masterContext);
                    masterRoot.Save();
                    changedParts.Add(PartPath(graph.Part));
                }
                var originalBackground = PptxBackgroundCodec.Read(masterCommon);
                var originalBackgroundHash = BackgroundSemanticHash(originalBackground);
                if (!binding.BackgroundSemanticSha256.Equals(originalBackgroundHash, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_master_source_background_mismatch",
                        $"Presentation master {masterIndex + 1} background does not match its source binding.",
                        PartPath(graph.Part));
                if (!BackgroundSemanticHash(target.Background).Equals(originalBackgroundHash, StringComparison.OrdinalIgnoreCase))
                {
                    if (!binding.BackgroundEditable || !PptxBackgroundCodec.Supports(masterCommon))
                        throw new CodecException("unsupported_presentation_edit", $"Presentation master {masterIndex + 1} background is preserved but not safely editable by this codec slice.", PartPath(graph.Part));
                    PptxBackgroundCodec.Apply(masterCommon, target.Background);
                    masterRoot.Save();
                    changedParts.Add(PartPath(graph.Part));
                }
                if (ApplyPlaceholders(masterShapeTree, graph.Id, target.Placeholders, masterContext, PartPath(graph.Part)))
                {
                    masterRoot.Save();
                    changedParts.Add(PartPath(graph.Part));
                }
                TrackContextChanges(graph.Part, masterContext, changedParts, addedRelationshipIds, addedPartPaths);
            }

            for (var layoutIndex = 0; layoutIndex < layoutGraph.Length; layoutIndex++)
            {
                var (master, graph) = layoutGraph[layoutIndex];
                var layoutRoot = graph.Part.SlideLayout ??
                    throw new CodecException("missing_slide_layout_root", $"Presentation layout {layoutIndex + 1} has no slide layout root.", PartPath(graph.Part));
                var layoutCommon = layoutRoot.CommonSlideData ??
                    throw new CodecException("missing_common_slide_data", $"Presentation layout {layoutIndex + 1} has no common slide data.", PartPath(graph.Part));
                var layoutShapeTree = layoutCommon.ShapeTree ??
                    throw new CodecException("missing_shape_tree", $"Presentation layout {layoutIndex + 1} has no shape tree.", PartPath(graph.Part));
                var target = envelope.Presentation.Layouts[layoutIndex];
                var binding = target.Source ?? throw new CodecException(
                    "missing_presentation_layout_binding",
                    $"Presentation layout {layoutIndex + 1} is missing its source binding.",
                    PartPath(graph.Part));
                var sourceName = layoutCommon.Name?.Value ?? $"Layout {graph.Index + 1}";
                if (target.Id != graph.Id || target.MasterId != master.Id || target.Name != sourceName || target.Type != LayoutTypeName(layoutRoot) ||
                    binding.LayoutIndex != graph.Index ||
                    !binding.PartPath.Equals(PartPath(graph.Part), StringComparison.OrdinalIgnoreCase) ||
                    !binding.RelationshipId.Equals(graph.RelationshipId, StringComparison.Ordinal) ||
                    !binding.LayoutXmlSha256.Equals(HashElement(layoutRoot), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_layout_binding_mismatch",
                        $"Presentation layout {layoutIndex + 1} does not match its hash-bound read-only source layout.",
                        PartPath(graph.Part));
                var layoutContext = new PptxPartContext(graph.Part, slideIdByPartPath, slidePartById, assetCatalog);
                var originalBackground = PptxBackgroundCodec.Read(layoutCommon);
                var originalBackgroundHash = BackgroundSemanticHash(originalBackground);
                if (!binding.BackgroundSemanticSha256.Equals(originalBackgroundHash, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_layout_source_background_mismatch",
                        $"Presentation layout {layoutIndex + 1} background does not match its source binding.",
                        PartPath(graph.Part));
                if (!BackgroundSemanticHash(target.Background).Equals(originalBackgroundHash, StringComparison.OrdinalIgnoreCase))
                {
                    if (!binding.BackgroundEditable || !PptxBackgroundCodec.Supports(layoutCommon))
                        throw new CodecException("unsupported_presentation_edit", $"Presentation layout {layoutIndex + 1} background is preserved but not safely editable by this codec slice.", PartPath(graph.Part));
                    PptxBackgroundCodec.Apply(layoutCommon, target.Background);
                    layoutRoot.Save();
                    changedParts.Add(PartPath(graph.Part));
                }
                if (ApplyPlaceholders(layoutShapeTree, graph.Id, target.Placeholders, layoutContext, PartPath(graph.Part)))
                {
                    layoutRoot.Save();
                    changedParts.Add(PartPath(graph.Part));
                }
                TrackContextChanges(graph.Part, layoutContext, changedParts, addedRelationshipIds, addedPartPaths);
            }

            ulong semanticItems = 0;
            for (var targetPosition = 0; targetPosition < retainedTargets.Length; targetPosition++)
            {
                var targetSlide = retainedTargets[targetPosition];
                var slideIndex = targetSlide.TargetIndex;
                var relationshipId = targetSlide.Source.RelationshipId;
                var slidePart = targetSlide.Source.Part;
                var slideRoot = slidePart.Slide ??
                    throw new CodecException("missing_slide_root", $"Presentation slide {slideIndex + 1} has no slide root.", PartPath(slidePart));
                var target = targetSlide.Target;
                var binding = target.Source ?? throw new CodecException(
                    "missing_presentation_slide_binding",
                    $"Presentation slide {slideIndex + 1} is missing its source binding.",
                    "ppt/presentation.xml");
                if (binding.SlideIndex != targetSlide.Source.Index ||
                    !binding.PartPath.Equals(PartPath(slidePart), StringComparison.OrdinalIgnoreCase) ||
                    !binding.RelationshipId.Equals(relationshipId, StringComparison.Ordinal) ||
                    !binding.SlideXmlSha256.Equals(HashElement(slideRoot), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_slide_binding_mismatch",
                        $"Presentation slide {slideIndex + 1} does not match its hash-bound source slide.",
                        PartPath(slidePart));
                var sourceLayoutPart = slidePart.SlideLayoutPart;
                var sourceLayoutId = sourceLayoutPart is null ? string.Empty :
                    layoutIdByPartPath.GetValueOrDefault(PartPath(sourceLayoutPart)) ??
                    throw new CodecException("unresolved_slide_layout_binding", $"Presentation slide {slideIndex + 1} references a layout outside the master graph.", PartPath(slidePart));
                var sourceLayoutRelationshipId = sourceLayoutPart is null ? string.Empty : slidePart.GetIdOfPart(sourceLayoutPart);
                if (target.LayoutId != sourceLayoutId || binding.LayoutRelationshipId != sourceLayoutRelationshipId)
                    throw new CodecException(
                        "presentation_slide_layout_binding_changed",
                        $"Source-preserving PPTX export cannot change slide {slideIndex + 1}'s layout binding.",
                        PartPath(slidePart));

                var slideCommon = slideRoot.CommonSlideData ??
                    throw new CodecException("missing_common_slide_data", $"Presentation slide {slideIndex + 1} has no common slide data.", PartPath(slidePart));
                var shapeTree = slideCommon.ShapeTree ??
                    throw new CodecException("missing_shape_tree", $"Presentation slide {slideIndex + 1} has no shape tree.", PartPath(slidePart));
                var originalBackground = PptxBackgroundCodec.Read(slideCommon);
                var originalBackgroundHash = BackgroundSemanticHash(originalBackground);
                if (!binding.BackgroundSemanticSha256.Equals(originalBackgroundHash, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_slide_source_background_mismatch",
                        $"Presentation slide {slideIndex + 1} background does not match its source binding.",
                        PartPath(slidePart));
                var changed = false;
                var sourceName = slideCommon.Name?.Value ?? $"Slide {targetSlide.Source.Index + 1}";
                if (!string.Equals(target.Name, sourceName, StringComparison.Ordinal))
                {
                    // This is deliberately the only source-bound slide metadata
                    // mutation: p:cSld/@name belongs to the existing SlidePart
                    // and does not alter its relationship graph or shape tree.
                    slideCommon.Name = target.Name;
                    changed = true;
                }
                if (!BackgroundSemanticHash(target.Background).Equals(originalBackgroundHash, StringComparison.OrdinalIgnoreCase))
                {
                    if (!binding.BackgroundEditable || !PptxBackgroundCodec.Supports(slideCommon))
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            $"Presentation slide {slideIndex + 1} background is preserved but not safely editable by this codec slice.",
                            PartPath(slidePart));
                    PptxBackgroundCodec.Apply(slideCommon, target.Background);
                    changed = true;
                }
                var sourceElements = ShapeElements(shapeTree);
                var elementIdsByNativeId = NativeElementIds(sourceElements, target.Id);
                var nativeIdsByElementId = elementIdsByNativeId.ToDictionary(item => item.Value, item => item.Key, StringComparer.Ordinal);
                if (sourceElements.Length != target.Elements.Count)
                    throw new CodecException(
                        "presentation_element_topology_changed",
                        $"Source-preserving PPTX export requires slide {slideIndex + 1}'s original {sourceElements.Length}-element topology; the artifact contains {target.Elements.Count} elements.",
                        PartPath(slidePart));

                var slideContext = new PptxPartContext(slidePart, slideIdByPartPath, slidePartById, assetCatalog);
                for (var elementIndex = 0; elementIndex < sourceElements.Length; elementIndex++)
                {
                    semanticItems++;
                    if (semanticItems > limits.MaxCells)
                        throw new CodecException("presentation_item_budget_exceeded", $"PPTX presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).", PartPath(slidePart));
                    var sourceElement = sourceElements[elementIndex];
                    var requested = target.Elements[elementIndex];
                    var elementBinding = requested.Source ?? throw new CodecException(
                        "missing_presentation_element_binding",
                        $"Presentation slide {slideIndex + 1} element {elementIndex + 1} is missing its source binding.",
                        PartPath(slidePart));
                    if (elementBinding.ShapeTreeIndex != elementIndex ||
                        !elementBinding.ElementSha256.Equals(HashElement(sourceElement), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "presentation_element_binding_mismatch",
                            $"Presentation slide {slideIndex + 1} element {elementIndex + 1} does not match its source element.",
                            PartPath(slidePart));
                    var original = ReadElement(sourceElement, slideIndex, elementIndex, slideContext, nativeObjects, elementIdsByNativeId);
                    if (original.ContentCase == PresentationElement.ContentOneofCase.Table)
                    {
                        semanticItems += checked((ulong)original.Table.Rows.Sum(row => row.Cells.Count));
                        if (semanticItems > limits.MaxCells)
                            throw new CodecException("presentation_item_budget_exceeded", $"PPTX presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).", PartPath(slidePart));
                    }
                    if (elementBinding.Editable != original.Source.Editable ||
                        elementBinding.TextEditable != original.Source.TextEditable)
                        throw new CodecException(
                            "presentation_element_binding_mismatch",
                            $"Presentation slide {slideIndex + 1} element {elementIndex + 1} changed its source editability contract.",
                            PartPath(slidePart));
                    if (!SemanticHash(original).Equals(elementBinding.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "presentation_source_semantics_mismatch",
                            $"Presentation slide {slideIndex + 1} element {elementIndex + 1} source semantics do not match its binding.",
                            PartPath(slidePart));
                    PptxOleWorkbookReplacement? oleWorkbookReplacement = null;
                    if (original.ContentCase == PresentationElement.ContentOneofCase.Opaque &&
                        requested.ContentCase == PresentationElement.ContentOneofCase.Opaque &&
                        PptxNativeObjectCatalog.SupportsPlacementEditing(sourceElement))
                    {
                        ValidateNativeObjectRequest(original, requested);
                        oleWorkbookReplacement = PptxOleWorkbookCodec.PrepareReplacement(original.Opaque, requested.Opaque, assetCatalog, limits);
                    }
                    if (SemanticHash(requested).Equals(elementBinding.SemanticSha256, StringComparison.OrdinalIgnoreCase)) continue;
                    if (!elementBinding.Editable)
                    {
                        if (elementBinding.TextEditable &&
                            sourceElement is P.Shape sourcePlaceholder &&
                            requested.ContentCase == PresentationElement.ContentOneofCase.Shape &&
                            PptxPlaceholderCodec.SupportsSlideTextEditing(sourcePlaceholder))
                        {
                            PptxPlaceholderCodec.ApplySlideText(sourcePlaceholder, original, requested, slideContext);
                            changed = true;
                            continue;
                        }
                        throw UnsupportedPresentationEdit(slideIndex, elementIndex, slidePart);
                    }
                    if (sourceElement is P.Shape sourceShape &&
                        requested.ContentCase == PresentationElement.ContentOneofCase.Shape &&
                        IsSimpleShape(sourceShape))
                    {
                        ApplyShape(sourceShape, requested, slideContext);
                        changed = true;
                    }
                    else if (sourceElement is P.Picture sourcePicture &&
                             requested.ContentCase == PresentationElement.ContentOneofCase.Image &&
                             PptxPictureCodec.TryRead(sourcePicture, slideContext, out _))
                    {
                        PptxPictureCodec.Apply(sourcePicture, requested, slideContext);
                        changed = true;
                    }
                    else if (sourceElement is P.GraphicFrame sourceTable &&
                             requested.ContentCase == PresentationElement.ContentOneofCase.Table &&
                             PptxTableCodec.TryRead(sourceTable, out _))
                    {
                        PptxTableCodec.Apply(sourceTable, requested);
                        changed = true;
                    }
                    else if (sourceElement is P.ConnectionShape sourceConnector &&
                             requested.ContentCase == PresentationElement.ContentOneofCase.Connector &&
                             TryReadConnector(sourceConnector, elementIdsByNativeId, out _))
                    {
                        ApplyConnector(sourceConnector, requested, nativeIdsByElementId);
                        changed = true;
                    }
                    else if (sourceElement is P.GraphicFrame sourceChart &&
                             requested.ContentCase == PresentationElement.ContentOneofCase.Chart &&
                             PptxChartCodec.TryRead(sourceChart, slideContext, out _, out var chartEditable) && chartEditable)
                    {
                        var replacement = PptxChartCodec.Apply(sourceChart, requested, slideContext);
                        changedParts.Add(replacement.PartPath);
                        replacedOpaquePartHashes.Add(replacement.PartPath, replacement.Sha256);
                        changed = true;
                    }
                    else if (sourceElement is P.GroupShape sourceGroup &&
                             requested.ContentCase == PresentationElement.ContentOneofCase.Group &&
                             original.ContentCase == PresentationElement.ContentOneofCase.Group &&
                             TryReadGroup(sourceGroup, original.Id, slideContext, elementIdsByNativeId, out _))
                    {
                        if (ApplyGroup(sourceGroup, original, requested, slideContext, elementIdsByNativeId, nativeIdsByElementId, changedParts, replacedOpaquePartHashes, slideIndex, $"element {elementIndex + 1}"))
                            changed = true;
                    }
                    else if (requested.ContentCase == PresentationElement.ContentOneofCase.Opaque &&
                             PptxNativeObjectCatalog.SupportsPlacementEditing(sourceElement))
                    {
                        if (oleWorkbookReplacement is not null)
                        {
                            PptxOleWorkbookCodec.Apply(slidePart, sourceElement, original.Opaque.OleWorkbook, oleWorkbookReplacement);
                            changedParts.Add(oleWorkbookReplacement.PartPath);
                            replacedOpaquePartHashes.Add(oleWorkbookReplacement.PartPath, oleWorkbookReplacement.Sha256);
                        }
                        if (NativePlacementChanged(original, requested))
                        {
                            ApplyNativePlacement(sourceElement, requested);
                            changed = true;
                        }
                    }
                    else
                    {
                        throw UnsupportedPresentationEdit(slideIndex, elementIndex, slidePart);
                    }
                }
                if (changed)
                {
                    slideRoot.Save();
                    changedParts.Add(PartPath(slidePart));
                }
                if (PptxSpeakerNotesCodec.ApplySourceBound(slidePart, target.SpeakerNotes, slideIndex) is { } notesChange)
                {
                    changedParts.Add(notesChange.PartPath);
                    replacedOpaquePartHashes.Add(notesChange.PartPath, notesChange.Sha256);
                }
                if (PptxModernCommentsCodec.ApplySourceBound(
                        presentationPart,
                        targetSlide.Source.SlideId,
                        slidePart,
                        sourceElements,
                        elementIdsByNativeId,
                        target,
                        slideIndex) is { } modernCommentsChange)
                {
                    changedParts.Add(modernCommentsChange.PartPath);
                    replacedOpaquePartHashes.Add(modernCommentsChange.PartPath, modernCommentsChange.Sha256);
                }
                TrackContextChanges(slidePart, slideContext, changedParts, addedRelationshipIds, addedPartPaths);
            }
        }

        var bytes = stream.ToArray();
        ValidateOutputBudget(bytes, limits);
        var retainedValidationErrorCount = ValidateOffice2021AgainstSource(sourceBytes, bytes);
        AssertPackagePartsUnchangedExcept(sourceBytes, bytes, changedParts);
        ValidatePreservedSlideElements(sourceBytes, bytes, envelope.Presentation, limits);
        ValidatePreservedMasterAndLayoutContent(sourceBytes, bytes, envelope.Presentation, limits);
        var outputOpaque = PackageGuards.ValidateAndCollectOpaque(bytes, limits, OpcPackageProfile.Pptx, includeSourcePackage: false);
        AssertOpaqueGraphMatchesWithModeledAdditions(
            envelope.OpaqueOpc,
            outputOpaque,
            addedRelationshipIds,
            addedPartPaths,
            replacedOpaquePartHashes,
            removedSourceSlidePartPaths);
        var diagnostics = new List<Diagnostic>();
        if (opaqueCount > 0)
            diagnostics.Add(CodecProtocol.Warning(
                "opaque_content_preserved",
                $"Preserved {opaqueCount} opaque OPC parts or relationships while updating modeled presentation content."));
        if (retainedValidationErrorCount > 0)
            diagnostics.Add(CodecProtocol.Warning(
                "source_openxml_validation_warnings_preserved",
                $"Preserved {retainedValidationErrorCount} pre-existing Office 2021 validation warning(s) from the source package; export introduced none."));
        return new PptxExportResult(bytes, diagnostics);
    }

    private static PresentationElement ReadElement(
        OpenXmlElement source,
        int slideIndex,
        int elementIndex,
        PptxPartContext slideContext,
        PptxNativeObjectCatalog? nativeObjects = null,
        IReadOnlyDictionary<uint, string>? elementIdsByNativeId = null)
        => ReadElement(source, $"presentation/slide/{slideIndex + 1}", elementIndex, slideContext, nativeObjects, elementIdsByNativeId);

    private static PresentationElement ReadElement(
        OpenXmlElement source,
        string ownerId,
        int elementIndex,
        PptxPartContext slideContext,
        PptxNativeObjectCatalog? nativeObjects = null,
        IReadOnlyDictionary<uint, string>? elementIdsByNativeId = null)
    {
        var element = new PresentationElement
        {
            Id = $"{ownerId}/element/{elementIndex + 1}",
            Name = ElementName(source, elementIndex),
        };
        var editable = source switch
        {
            P.Shape shape => IsSimpleShape(shape),
            P.Picture picture => PptxPictureCodec.TryRead(picture, slideContext, out _),
            P.GraphicFrame graphicFrame => PptxTableCodec.TryRead(graphicFrame, out _),
            P.ConnectionShape connector => TryReadConnector(connector, elementIdsByNativeId, out _),
            P.GroupShape group => TryReadGroup(group, element.Id, slideContext, elementIdsByNativeId, out _),
            _ => false,
        };
        if (source is P.GraphicFrame chartFrame && PptxChartCodec.TryRead(chartFrame, slideContext, out _, out var chartEditable)) editable = chartEditable;
        if (source is P.Shape sourceShape)
            element.Shape = ReadShape(sourceShape, slideContext);
        else if (source is P.Picture sourcePicture && PptxPictureCodec.TryRead(sourcePicture, slideContext, out var image))
            element.Image = image;
        else if (source is P.GraphicFrame sourceTable && PptxTableCodec.TryRead(sourceTable, out var table))
            element.Table = table;
        else if (source is P.ConnectionShape sourceConnector && TryReadConnector(sourceConnector, elementIdsByNativeId, out var connector))
            element.Connector = connector;
        else if (source is P.GraphicFrame sourceChart && PptxChartCodec.TryRead(sourceChart, slideContext, out var chart, out _))
            element.Chart = chart;
        else if (source is P.GroupShape sourceGroup && TryReadGroup(sourceGroup, element.Id, slideContext, elementIdsByNativeId, out var group))
            element.Group = group;
        else
        {
            var frame = ReadFrame(source);
            element.Opaque = new PresentationOpaqueElement
            {
                ElementName = source.LocalName,
                Text = DescendantText(source),
                RawXml = source.OuterXml,
                LeftEmu = frame.Left,
                TopEmu = frame.Top,
                WidthEmu = frame.Width,
                HeightEmu = frame.Height,
            };
            nativeObjects?.Populate(element.Opaque, source, PartPath(slideContext.Owner));
            editable = PptxNativeObjectCatalog.SupportsPlacementEditing(source);
        }
        element.Source = new PresentationElementSourceBinding
        {
            ShapeTreeIndex = checked((uint)elementIndex),
            ElementSha256 = HashElement(source),
            Editable = editable,
            TextEditable = source is P.Shape placeholderShape && PptxPlaceholderCodec.SupportsSlideTextEditing(placeholderShape),
        };
        element.Source.SemanticSha256 = SemanticHash(element);
        return element;
    }

    private static bool TryReadGroup(
        P.GroupShape source,
        string groupId,
        PptxPartContext slideContext,
        IReadOnlyDictionary<uint, string>? elementIdsByNativeId,
        out PresentationGroup group)
    {
        group = new PresentationGroup();
        var nonVisual = source.GetFirstChild<P.NonVisualGroupShapeProperties>();
        var properties = source.GetFirstChild<P.GroupShapeProperties>();
        var transform = properties?.GetFirstChild<A.TransformGroup>();
        if (nonVisual is null || properties is null || transform is null ||
            source.Elements<P.NonVisualGroupShapeProperties>().Count() != 1 ||
            source.Elements<P.GroupShapeProperties>().Count() != 1 ||
            nonVisual.ChildElements.Count != 3 ||
            nonVisual.ChildElements[0] is not P.NonVisualDrawingProperties drawing ||
            nonVisual.ChildElements[1] is not P.NonVisualGroupShapeDrawingProperties groupDrawing ||
            nonVisual.ChildElements[2] is not P.ApplicationNonVisualDrawingProperties application ||
            drawing.ChildElements.Count != 0 || groupDrawing.ChildElements.Count != 0 || application.ChildElements.Count != 0 ||
            !HasOnlyAttributes(drawing, "id", "name") || !HasOnlyAttributes(groupDrawing) || !HasOnlyAttributes(application) ||
            properties.ChildElements.Count != 1 || properties.FirstChild != transform || !HasOnlyAttributes(properties) ||
            !HasOnlyAttributes(transform) || transform.ChildElements.Count != 4 ||
            transform.ChildElements[0] is not A.Offset offset ||
            transform.ChildElements[1] is not A.Extents extents ||
            transform.ChildElements[2] is not A.ChildOffset childOffset ||
            transform.ChildElements[3] is not A.ChildExtents childExtents ||
            !HasOnlyAttributes(offset, "x", "y") || !HasOnlyAttributes(extents, "cx", "cy") ||
            !HasOnlyAttributes(childOffset, "x", "y") || !HasOnlyAttributes(childExtents, "cx", "cy") ||
            extents.Cx?.Value <= 0 || extents.Cy?.Value <= 0 || childExtents.Cx?.Value <= 0 || childExtents.Cy?.Value <= 0 ||
            offset.X?.Value < 0 || offset.Y?.Value < 0)
            return false;

        group.LeftEmu = offset.X?.Value ?? 0;
        group.TopEmu = offset.Y?.Value ?? 0;
        group.WidthEmu = extents.Cx?.Value ?? 0;
        group.HeightEmu = extents.Cy?.Value ?? 0;
        group.ChildLeftEmu = childOffset.X?.Value ?? 0;
        group.ChildTopEmu = childOffset.Y?.Value ?? 0;
        group.ChildWidthEmu = childExtents.Cx?.Value ?? 0;
        group.ChildHeightEmu = childExtents.Cy?.Value ?? 0;
        var children = GroupElements(source);
        if (children.Length == 0) return false;
        for (var index = 0; index < children.Length; index++)
        {
            var child = ReadElement(children[index], groupId, index, slideContext, elementIdsByNativeId: elementIdsByNativeId);
            if (child.ContentCase is PresentationElement.ContentOneofCase.Opaque or PresentationElement.ContentOneofCase.None || child.Source?.Editable != true)
                return false;
            group.Children.Add(child);
        }
        return true;
    }

    private static PresentationShape ReadShape(P.Shape shape, PptxPartContext slideContext)
    {
        var frame = ReadFrame(shape);
        var properties = shape.ShapeProperties;
        var textBody = PptxTextCodec.Read(shape.TextBody, slideContext);
        var placeholder = PptxPlaceholderCodec.ReadIdentity(shape);
        var transform = properties?.Transform2D;
        var result = new PresentationShape
        {
            Geometry = Geometry(shape),
            LeftEmu = frame.Left,
            TopEmu = frame.Top,
            WidthEmu = frame.Width,
            HeightEmu = frame.Height,
            Text = PptxTextCodec.Flatten(textBody),
            TextBody = textBody,
            FillRgb = PptxColor.SolidRgb(properties?.GetFirstChild<A.SolidFill>()),
            LineRgb = PptxColor.SolidRgb(properties?.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>()),
            LineWidthEmu = properties?.GetFirstChild<A.Outline>()?.Width?.Value ?? 0,
            Placeholder = placeholder,
            DirectFrame = placeholder is null ? null : PptxPlaceholderCodec.ReadDirectFrame(shape),
            Transform = placeholder is null && PptxShapeTransformCodec.Supports(transform)
                ? PptxShapeTransformCodec.Read(transform!)
                : null,
            Shadow = ReadShadow(properties),
        };
        if (shape.UseBackgroundFill?.HasValue == true)
            result.UseBackgroundFill = shape.UseBackgroundFill.Value;
        result.CustomPaths.Add(PptxCustomGeometryCodec.Read(properties?.GetFirstChild<A.CustomGeometry>()));
        return result;
    }

    private static bool IsSimpleShape(P.Shape shape)
    {
        if (shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>() is not null) return false;
        if (shape.ShapeStyle is not null) return false;
        var properties = shape.ShapeProperties;
        var transform = properties?.Transform2D;
        if (properties is null || properties.Elements<A.Transform2D>().Count() != 1 || !PptxShapeTransformCodec.Supports(transform)) return false;
        var geometry = Geometry(shape);
        if (geometry is not ("rect" or "ellipse" or "roundRect" or "textbox" or "custom")) return false;
        if (geometry == "custom" && !PptxCustomGeometryCodec.Supports(properties.GetFirstChild<A.CustomGeometry>())) return false;
        if (!SimpleFill(properties)) return false;
        var outline = properties.GetFirstChild<A.Outline>();
        if (outline is not null && !SimpleFill(outline)) return false;
        if (!SupportsShadow(properties)) return false;
        if (properties.ChildElements.Any(child => child is not A.Transform2D and not A.PresetGeometry and not A.CustomGeometry and not A.NoFill and not A.SolidFill and not A.Outline and not A.EffectList)) return false;
        return PptxTextCodec.SupportsEditing(shape.TextBody);
    }

    private static bool SimpleFill(OpenXmlCompositeElement element)
    {
        var fills = element.ChildElements.Where(child => child is A.NoFill or A.SolidFill).ToArray();
        if (fills.Length > 1) return false;
        if (fills.Length == 0 || fills[0] is A.NoFill) return true;
        var solid = (A.SolidFill)fills[0];
        return solid.ChildElements.Count == 1 && solid.FirstChild is A.RgbColorModelHex;
    }

    private static PresentationShadow? ReadShadow(P.ShapeProperties? properties)
    {
        if (!SupportsShadow(properties)) return null;
        var outer = properties?.GetFirstChild<A.EffectList>()?.GetFirstChild<A.OuterShadow>();
        if (outer?.GetFirstChild<A.RgbColorModelHex>() is not { } color) return null;
        return new PresentationShadow
        {
            ColorRgb = PptxColor.Normalize(color.Val?.Value ?? string.Empty),
            BlurRadiusEmu = outer.BlurRadius?.Value ?? 0L,
            DistanceEmu = outer.Distance?.Value ?? 0L,
            DirectionAngle60000 = outer.Direction?.Value ?? 0,
            OpacityThousandthPercent = checked((uint)(color.GetFirstChild<A.Alpha>()?.Val?.Value ?? 100_000)),
        };
    }

    private static bool SupportsShadow(P.ShapeProperties? properties)
    {
        var lists = properties?.Elements<A.EffectList>().ToArray() ?? [];
        if (lists.Length == 0) return true;
        if (lists.Length != 1 || lists[0].ChildElements.Count != 1 || lists[0].FirstChild is not A.OuterShadow outer ||
            !HasOnlyAttributes(outer, "blurRad", "dist", "dir") || outer.BlurRadius?.Value is < 0 || outer.Distance?.Value is < 0 ||
            outer.Direction?.Value is < 0 or >= 21_600_000 || outer.ChildElements.Count != 1 || outer.FirstChild is not A.RgbColorModelHex color ||
            color.Val?.Value is not { Length: 6 } rgb || !rgb.All(Uri.IsHexDigit) || !HasOnlyAttributes(color, "val")) return false;
        var alphas = color.Elements<A.Alpha>().ToArray();
        return color.ChildElements.Count == alphas.Length && alphas.Length <= 1 &&
               (alphas.Length == 0 || alphas[0].Val?.Value is >= 0 and <= 100_000 && HasOnlyAttributes(alphas[0], "val"));
    }

    private static void ApplyShadow(P.ShapeProperties properties, PresentationShadow? shadow)
    {
        properties.GetFirstChild<A.EffectList>()?.Remove();
        if (shadow is null) return;
        var color = new A.RgbColorModelHex { Val = PptxColor.Normalize(shadow.ColorRgb) };
        color.Append(new A.Alpha { Val = checked((int)shadow.OpacityThousandthPercent) });
        var outer = new A.OuterShadow(color)
        {
            BlurRadius = shadow.BlurRadiusEmu,
            Distance = shadow.DistanceEmu,
            Direction = shadow.DirectionAngle60000,
        };
        properties.Append(new A.EffectList(outer));
    }

    private static void ValidateShadow(PresentationShadow? shadow, string elementId)
    {
        if (shadow is null) return;
        PptxColor.Normalize(shadow.ColorRgb);
        if (shadow.BlurRadiusEmu < 0 || shadow.DistanceEmu < 0 || shadow.DirectionAngle60000 is < 0 or >= 21_600_000 || shadow.OpacityThousandthPercent > 100_000)
            throw new CodecException("invalid_presentation_shadow", $"Presentation shape {elementId} has invalid shadow geometry or opacity.");
    }

    private static void ApplyShape(P.Shape shape, PresentationElement source, PptxPartContext slideContext)
    {
        var semantic = source.Shape;
        var sourceHasBackgroundFill = shape.UseBackgroundFill?.HasValue == true;
        if (sourceHasBackgroundFill != semantic.HasUseBackgroundFill ||
            sourceHasBackgroundFill && shape.UseBackgroundFill!.Value != semantic.UseBackgroundFill)
            throw new CodecException(
                "unsupported_presentation_edit",
                $"Presentation shape {source.Id} cannot change its source-bound useBgFill attribute.");
        var properties = shape.ShapeProperties ??= new P.ShapeProperties();
        var transform = properties.Transform2D ??= new A.Transform2D();
        var offset = transform.Offset ??= new A.Offset();
        offset.X = semantic.LeftEmu;
        offset.Y = semantic.TopEmu;
        var extents = transform.Extents ??= new A.Extents();
        extents.Cx = semantic.WidthEmu;
        extents.Cy = semantic.HeightEmu;
        PptxShapeTransformCodec.Apply(transform, semantic.Transform);
        PptxCustomGeometryCodec.Apply(properties, semantic);
        if (shape.NonVisualShapeProperties?.NonVisualShapeDrawingProperties is { } drawingProperties)
            drawingProperties.TextBox = semantic.Geometry == "textbox" ? true : null;
        if (!FillMatches(properties, semantic.FillRgb)) ReplaceFill(properties, semantic.FillRgb);
        var outline = properties.GetFirstChild<A.Outline>();
        if (outline is null && (semantic.LineWidthEmu > 0 || !string.IsNullOrWhiteSpace(semantic.LineRgb)))
        {
            outline = new A.Outline();
            properties.Append(outline);
        }
        if (outline is not null)
        {
            outline.Width = checked((int)semantic.LineWidthEmu);
            if (!FillMatches(outline, semantic.LineRgb)) ReplaceFill(outline, semantic.LineRgb);
        }
        if (shape.NonVisualShapeProperties?.NonVisualDrawingProperties is { } nonVisual)
            nonVisual.Name = source.Name;
        ApplyShadow(properties, semantic.Shadow);
        PptxTextCodec.Apply(shape, semantic, slideContext);
    }

    private static CodecException UnsupportedPresentationEdit(int slideIndex, int elementIndex, OpenXmlPart slidePart) => new(
        "unsupported_presentation_edit",
        $"Presentation slide {slideIndex + 1} element {elementIndex + 1} is preserved but not safely editable by this codec slice.",
        PartPath(slidePart));

    private static void ValidateNativeObjectRequest(PresentationElement original, PresentationElement requested)
    {
        var allowed = original.Clone();
        allowed.Name = requested.Name;
        allowed.Opaque.LeftEmu = requested.Opaque.LeftEmu;
        allowed.Opaque.TopEmu = requested.Opaque.TopEmu;
        allowed.Opaque.WidthEmu = requested.Opaque.WidthEmu;
        allowed.Opaque.HeightEmu = requested.Opaque.HeightEmu;
        if (allowed.Opaque.OleWorkbook is not null && requested.Opaque.OleWorkbook is not null)
            allowed.Opaque.OleWorkbook.ReplacementAssetId = requested.Opaque.OleWorkbook.ReplacementAssetId;
        // Source binding equality is checked against the actual source above;
        // reuse the caller's equivalent instance to keep protobuf equality
        // focused on the semantic payload.
        allowed.Source = requested.Source.Clone();
        if (!allowed.Equals(requested))
            throw new CodecException(
                "unsupported_presentation_edit",
                $"Presentation native object {requested.Id} may edit only its name, outer frame, and an explicitly recognized OLE workbook payload.");
    }

    private static bool NativePlacementChanged(PresentationElement original, PresentationElement requested) =>
        original.Name != requested.Name ||
        original.Opaque.LeftEmu != requested.Opaque.LeftEmu ||
        original.Opaque.TopEmu != requested.Opaque.TopEmu ||
        original.Opaque.WidthEmu != requested.Opaque.WidthEmu ||
        original.Opaque.HeightEmu != requested.Opaque.HeightEmu;

    private static void ApplyNativePlacement(OpenXmlElement source, PresentationElement requested)
    {
        var frame = requested.Opaque;
        if (source is P.GraphicFrame graphicFrame)
        {
            if (graphicFrame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties is { } nonVisual)
                nonVisual.Name = requested.Name;
            SetFrame(graphicFrame.Transform!, frame);
            if (PptxNativeObjectCatalog.Classify(source) == "oleObject")
            {
                // PowerPoint stores a second transform on the OLE preview
                // picture. Keep it derived from the outer frame so Office and
                // fallback renderers agree after a move/resize.
                var previewTransform = graphicFrame.Descendants<A.Transform2D>().FirstOrDefault();
                if (previewTransform is not null) SetFrame(previewTransform, frame);
            }
            return;
        }
        if (source is P.GroupShape group)
        {
            if (group.GetFirstChild<P.NonVisualGroupShapeProperties>()?.NonVisualDrawingProperties is { } nonVisual)
                nonVisual.Name = requested.Name;
            SetFrame(group.GetFirstChild<P.GroupShapeProperties>()!.GetFirstChild<A.TransformGroup>()!, frame);
            return;
        }
        throw new CodecException("unsupported_presentation_edit", $"Presentation native object {requested.Id} has no supported placement owner.");
    }

    private static void SetFrame(P.Transform transform, PresentationOpaqueElement frame)
    {
        transform.Offset!.X = frame.LeftEmu;
        transform.Offset.Y = frame.TopEmu;
        transform.Extents!.Cx = frame.WidthEmu;
        transform.Extents.Cy = frame.HeightEmu;
    }

    private static void SetFrame(A.Transform2D transform, PresentationOpaqueElement frame)
    {
        transform.Offset ??= new A.Offset();
        transform.Extents ??= new A.Extents();
        transform.Offset.X = frame.LeftEmu;
        transform.Offset.Y = frame.TopEmu;
        transform.Extents.Cx = frame.WidthEmu;
        transform.Extents.Cy = frame.HeightEmu;
    }

    private static void SetFrame(A.TransformGroup transform, PresentationOpaqueElement frame)
    {
        transform.Offset!.X = frame.LeftEmu;
        transform.Offset.Y = frame.TopEmu;
        transform.Extents!.Cx = frame.WidthEmu;
        transform.Extents.Cy = frame.HeightEmu;
    }

    private static void ReplaceFill(OpenXmlCompositeElement parent, string rgb)
    {
        foreach (var child in parent.ChildElements.Where(child => child is A.NoFill or A.SolidFill).ToArray()) child.Remove();
        OpenXmlElement fill = string.IsNullOrWhiteSpace(rgb)
            ? new A.NoFill()
            : new A.SolidFill(new A.RgbColorModelHex { Val = PptxColor.Normalize(rgb) });
        var reference = parent.ChildElements.FirstOrDefault(child => child is A.Outline || child.LocalName is "effectLst" or "effectDag" or "scene3d" or "sp3d");
        if (reference is null) parent.Append(fill);
        else parent.InsertBefore(fill, reference);
    }

    private static bool FillMatches(OpenXmlCompositeElement parent, string rgb)
    {
        var requested = string.IsNullOrWhiteSpace(rgb) ? string.Empty : PptxColor.Normalize(rgb);
        if (parent.GetFirstChild<A.NoFill>() is not null) return requested.Length == 0;
        var solid = PptxColor.SolidRgb(parent.GetFirstChild<A.SolidFill>());
        if (solid.Length > 0) return requested.Equals(solid, StringComparison.OrdinalIgnoreCase);
        return requested.Length == 0 && !parent.ChildElements.Any(child => child.LocalName.EndsWith("Fill", StringComparison.Ordinal));
    }

    private static void BuildPresentation(PresentationDocument package, PresentationArtifact artifact, PptxAssetCatalog assetCatalog)
    {
        if (artifact.Masters.Count > 1)
            throw new CodecException(
                "unsupported_presentation_features",
                "New PPTX authoring supports one canonical master. Multiple slide-master graphs require a validated source package.");
        var authoredMaster = artifact.Masters.FirstOrDefault();
        var canonicalMasterId = authoredMaster?.Id ?? "__openchestnut/default-master";
        var presentationPart = package.AddPresentationPart();
        var masterPart = presentationPart.AddNewPart<SlideMasterPart>("rIdMaster1");
        var themePart = masterPart.AddNewPart<ThemePart>("rIdTheme1");

        themePart.Theme = BasicTheme();
        var sourceLayouts = artifact.Layouts.ToList();
        PresentationLayout? fallbackLayout = null;
        if (sourceLayouts.Count == 0 || artifact.Slides.Any(slide => string.IsNullOrWhiteSpace(slide.LayoutId)))
        {
            var fallbackId = "__openchestnut/default-layout";
            while (sourceLayouts.Any(layout => layout.Id.Equals(fallbackId, StringComparison.Ordinal))) fallbackId += "-";
            fallbackLayout = new PresentationLayout
            {
                Id = fallbackId,
                Name = "Blank",
                MasterId = canonicalMasterId,
                Type = "blank",
            };
            sourceLayouts.Insert(0, fallbackLayout);
        }

        var layoutIdList = new P.SlideLayoutIdList();
        var layoutEntries = new List<(PresentationLayout Layout, SlideLayoutPart Part)>();
        for (var layoutIndex = 0; layoutIndex < sourceLayouts.Count; layoutIndex++)
        {
            var sourceLayout = sourceLayouts[layoutIndex];
            var relationshipId = $"rIdLayout{layoutIndex + 1}";
            var layoutPart = masterPart.AddNewPart<SlideLayoutPart>(relationshipId);
            layoutPart.AddPart(masterPart, "rIdMaster1");
            var layoutRoot = new P.SlideLayout(
                new P.CommonSlideData(BasicShapeTree()) { Name = string.IsNullOrWhiteSpace(sourceLayout.Name) ? "Blank" : sourceLayout.Name },
                new P.ColorMapOverride(new A.MasterColorMapping()))
            { Preserve = true };
            layoutRoot.SetAttribute(new OpenXmlAttribute("type", string.Empty, SourceFreeLayoutType(sourceLayout)));
            layoutPart.SlideLayout = layoutRoot;
            layoutIdList.Append(new P.SlideLayoutId
            {
                Id = checked(2_147_483_649U + (uint)layoutIndex),
                RelationshipId = relationshipId,
            });
            layoutEntries.Add((sourceLayout, layoutPart));
        }
        var layoutPartById = layoutEntries.ToDictionary(entry => entry.Layout.Id, entry => entry.Part, StringComparer.Ordinal);
        var defaultLayoutPart = fallbackLayout is not null
            ? layoutPartById[fallbackLayout.Id]
            : layoutEntries[0].Part;
        masterPart.SlideMaster = new P.SlideMaster(
            new P.CommonSlideData(BasicShapeTree()) { Name = string.IsNullOrWhiteSpace(authoredMaster?.Name) ? "Office Clean Room" : authoredMaster.Name },
            BasicColorMap(),
            layoutIdList,
            new P.TextStyles(new P.TitleStyle(), new P.BodyStyle(), new P.OtherStyle()));

        var slideIdList = new P.SlideIdList();
        var slideParts = new SlidePart[artifact.Slides.Count];
        for (var slideIndex = 0; slideIndex < artifact.Slides.Count; slideIndex++)
        {
            var source = artifact.Slides[slideIndex];
            var relationshipId = $"rIdSlide{slideIndex + 1}";
            var slidePart = presentationPart.AddNewPart<SlidePart>(relationshipId);
            slideParts[slideIndex] = slidePart;
            var layoutPart = string.IsNullOrWhiteSpace(source.LayoutId)
                ? defaultLayoutPart
                : layoutPartById.GetValueOrDefault(source.LayoutId) ??
                  throw new CodecException("invalid_presentation_layout", $"Presentation slide {source.Id} references missing layout {source.LayoutId}.");
            slidePart.AddPart(layoutPart, "rIdLayout1");
            slidePart.Slide = new P.Slide(
                new P.CommonSlideData(BasicShapeTree()) { Name = source.Name },
                new P.ColorMapOverride(new A.MasterColorMapping()));
            slideIdList.Append(new P.SlideId { Id = checked((uint)(256 + slideIndex)), RelationshipId = relationshipId });
        }
        var slideIdByPartPath = slideParts
            .Select((part, index) => (Path: PartPath(part), Id: artifact.Slides[index].Id))
            .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
        var slidePartById = slideParts
            .Select((part, index) => (Part: part, Id: artifact.Slides[index].Id))
            .ToDictionary(item => item.Id, item => item.Part, StringComparer.Ordinal);
        var masterContext = new PptxPartContext(masterPart, slideIdByPartPath, slidePartById, assetCatalog);
        PptxBackgroundCodec.Build(masterPart.SlideMaster.CommonSlideData!, authoredMaster?.Background);
        PptxMasterTextStylesCodec.Build(masterPart.SlideMaster, authoredMaster?.TextStyles, masterContext);
        var masterShapeTree = masterPart.SlideMaster.CommonSlideData!.ShapeTree!;
        foreach (var (placeholder, index) in (authoredMaster?.Placeholders ?? []).Select((placeholder, index) => (placeholder, index)))
            masterShapeTree.Append(PptxPlaceholderCodec.Build(placeholder, checked((uint)(index + 2)), masterContext));
        foreach (var (layout, layoutPart) in layoutEntries)
        {
            var layoutContext = new PptxPartContext(layoutPart, slideIdByPartPath, slidePartById, assetCatalog);
            var layoutCommon = layoutPart.SlideLayout!.CommonSlideData!;
            PptxBackgroundCodec.Build(layoutCommon, layout.Background);
            var layoutShapeTree = layoutCommon.ShapeTree!;
            foreach (var (placeholder, index) in layout.Placeholders.Select((placeholder, index) => (placeholder, index)))
                layoutShapeTree.Append(PptxPlaceholderCodec.Build(placeholder, checked((uint)(index + 2)), layoutContext));
        }
        for (var slideIndex = 0; slideIndex < artifact.Slides.Count; slideIndex++)
        {
            var source = artifact.Slides[slideIndex];
            var slidePart = slideParts[slideIndex];
            var slideCommon = slidePart.Slide!.CommonSlideData!;
            PptxBackgroundCodec.Build(slideCommon, source.Background);
            var shapeTree = slideCommon.ShapeTree!;
            var slideContext = new PptxPartContext(slidePart, slideIdByPartPath, slidePartById, assetCatalog);
            var flattenedElements = FlattenPresentationElements(source.Elements).ToArray();
            var nativeIdsByElementId = flattenedElements.Select((element, index) => (element.Id, NativeId: checked((uint)(index + 2))))
                .ToDictionary(item => item.Id, item => item.NativeId, StringComparer.Ordinal);
            foreach (var element in source.Elements)
                shapeTree.Append(BuildElement(element, nativeIdsByElementId, slideContext, slidePart));
            slidePart.Slide.Save();
        }
        var notesMasterRelationshipId = PptxSpeakerNotesCodec.BuildSourceFree(presentationPart, themePart, slideParts, artifact.Slides);
        var presentationRoot = new P.Presentation();
        presentationRoot.Append(new P.SlideMasterIdList(new P.SlideMasterId { Id = 2_147_483_648U, RelationshipId = "rIdMaster1" }));
        if (notesMasterRelationshipId is not null)
            presentationRoot.Append(new P.NotesMasterIdList(new P.NotesMasterId { Id = notesMasterRelationshipId }));
        presentationRoot.Append(
            slideIdList,
            new P.SlideSize
            {
                Cx = checked((int)(artifact.SlideWidthEmu > 0 ? artifact.SlideWidthEmu : DefaultSlideWidthEmu)),
                Cy = checked((int)(artifact.SlideHeightEmu > 0 ? artifact.SlideHeightEmu : DefaultSlideHeightEmu)),
            },
            new P.NotesSize { Cx = 6_858_000L, Cy = 9_144_000L },
            new P.DefaultTextStyle());
        presentationPart.Presentation = presentationRoot;
        PptxLegacyCommentsCodec.BuildSourceFree(presentationPart, slideParts, artifact.Slides);
        PptxModernCommentsCodec.BuildSourceFree(presentationPart, slideIdList.Elements<P.SlideId>().ToArray(), slideParts, artifact.Slides);
        themePart.Theme.Save();
        foreach (var (_, layoutPart) in layoutEntries) layoutPart.SlideLayout!.Save();
        masterPart.SlideMaster.Save();
        presentationPart.Presentation.Save();
    }

    private static string SourceFreeLayoutType(PresentationLayout layout)
    {
        var type = string.IsNullOrWhiteSpace(layout.Type) ? "blank" : layout.Type;
        return type switch
        {
            "blank" or "title" or "titleOnly" or "obj" => type,
            _ => throw new CodecException(
                "unsupported_presentation_features",
                $"Source-free presentation layout {layout.Id} uses unsupported type {type}. Use blank, title, titleOnly, or obj."),
        };
    }

    private static void ValidateSourceFreeTextPlaceholder(PresentationPlaceholder placeholder, string ownerId)
    {
        ValidateSourceFreeTextPlaceholderType(placeholder.Type, placeholder.Id);
        if (placeholder.DirectFrame is null)
            throw new CodecException("invalid_presentation_placeholder", $"Source-free {ownerId} placeholder {placeholder.Id} requires a direct frame.");
    }

    private static void ValidateSourceFreeTextPlaceholderIdentity(PresentationPlaceholderIdentity placeholder, string elementId) =>
        ValidateSourceFreeTextPlaceholderType(placeholder.Type, elementId);

    private static void ValidateSourceFreeTextPlaceholderType(string type, string ownerId)
    {
        if (type is "title" or "body" or "ctrTitle" or "subTitle") return;
        throw new CodecException(
            "unsupported_presentation_features",
            $"Source-free presentation placeholder {ownerId} uses {type}; only title, body, ctrTitle, and subTitle text placeholders are supported.");
    }

    private static IEnumerable<PresentationElement> FlattenPresentationElements(IEnumerable<PresentationElement> elements)
    {
        foreach (var element in elements)
        {
            yield return element;
            if (element.ContentCase == PresentationElement.ContentOneofCase.Group)
                foreach (var child in FlattenPresentationElements(element.Group.Children)) yield return child;
        }
    }

    private static OpenXmlElement BuildElement(
        PresentationElement element,
        IReadOnlyDictionary<string, uint> nativeIdsByElementId,
        PptxPartContext slideContext,
        SlidePart slidePart) => element.ContentCase switch
        {
            PresentationElement.ContentOneofCase.Shape => BuildShape(element, nativeIdsByElementId[element.Id], slideContext),
            PresentationElement.ContentOneofCase.Image => PptxPictureCodec.Build(element, nativeIdsByElementId[element.Id], slideContext),
            PresentationElement.ContentOneofCase.Table => PptxTableCodec.Build(element, nativeIdsByElementId[element.Id]),
            PresentationElement.ContentOneofCase.Connector => BuildConnector(element, nativeIdsByElementId[element.Id], nativeIdsByElementId),
            PresentationElement.ContentOneofCase.Chart => PptxChartCodec.Build(element, nativeIdsByElementId[element.Id], slidePart),
            PresentationElement.ContentOneofCase.Group => BuildGroup(element, nativeIdsByElementId, slideContext, slidePart),
            _ => throw new CodecException("unsupported_presentation_element", $"Opaque presentation element {element.Id} requires its validated source package and cannot be authored from scratch."),
        };

    private static P.GroupShape BuildGroup(
        PresentationElement element,
        IReadOnlyDictionary<string, uint> nativeIdsByElementId,
        PptxPartContext slideContext,
        SlidePart slidePart)
    {
        var group = element.Group;
        var output = new P.GroupShape(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = nativeIdsByElementId[element.Id], Name = element.Name },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(new A.TransformGroup(
                new A.Offset { X = group.LeftEmu, Y = group.TopEmu },
                new A.Extents { Cx = group.WidthEmu, Cy = group.HeightEmu },
                new A.ChildOffset { X = group.ChildLeftEmu, Y = group.ChildTopEmu },
                new A.ChildExtents { Cx = group.ChildWidthEmu, Cy = group.ChildHeightEmu })));
        foreach (var child in group.Children)
            output.Append(BuildElement(child, nativeIdsByElementId, slideContext, slidePart));
        return output;
    }

    private static P.Shape BuildShape(PresentationElement source, uint nativeId, PptxPartContext slideContext)
    {
        var semantic = source.Shape;
        var directFrame = semantic.DirectFrame;
        var transform = new A.Transform2D(
            new A.Offset { X = directFrame?.LeftEmu ?? semantic.LeftEmu, Y = directFrame?.TopEmu ?? semantic.TopEmu },
            new A.Extents { Cx = directFrame?.WidthEmu ?? semantic.WidthEmu, Cy = directFrame?.HeightEmu ?? semantic.HeightEmu });
        if (directFrame is not null)
        {
            transform.Rotation = directFrame.HasRotationAngle60000 ? directFrame.RotationAngle60000 : null;
            transform.HorizontalFlip = directFrame.HasFlipHorizontal ? directFrame.FlipHorizontal : null;
            transform.VerticalFlip = directFrame.HasFlipVertical ? directFrame.FlipVertical : null;
        }
        else
        {
            PptxShapeTransformCodec.Apply(transform, semantic.Transform);
        }
        var properties = new P.ShapeProperties(transform);
        PptxCustomGeometryCodec.Apply(properties, semantic);
        properties.Append(string.IsNullOrWhiteSpace(semantic.FillRgb)
            ? new A.NoFill()
            : new A.SolidFill(new A.RgbColorModelHex { Val = PptxColor.Normalize(semantic.FillRgb) }));
        var outline = new A.Outline { Width = checked((int)semantic.LineWidthEmu) };
        outline.Append(string.IsNullOrWhiteSpace(semantic.LineRgb)
            ? new A.NoFill()
            : new A.SolidFill(new A.RgbColorModelHex { Val = PptxColor.Normalize(semantic.LineRgb) }));
        properties.Append(outline);
        ApplyShadow(properties, semantic.Shadow);
        var applicationProperties = new P.ApplicationNonVisualDrawingProperties();
        if (semantic.Placeholder is not null)
        {
            var nativePlaceholder = new P.PlaceholderShape { Index = semantic.Placeholder.Index };
            nativePlaceholder.SetAttribute(new OpenXmlAttribute("type", string.Empty, semantic.Placeholder.Type));
            applicationProperties.Append(nativePlaceholder);
        }
        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = nativeId, Name = source.Name },
                new P.NonVisualShapeDrawingProperties { TextBox = semantic.Geometry == "textbox" ? true : null },
                applicationProperties),
            properties,
            PptxTextCodec.Build(semantic, slideContext));
    }

    private static bool TryReadConnector(P.ConnectionShape source, IReadOnlyDictionary<uint, string>? elementIdsByNativeId, out PresentationConnector connector)
    {
        connector = new PresentationConnector();
        var properties = source.ShapeProperties;
        var transform = properties?.Transform2D;
        var geometry = properties?.GetFirstChild<A.PresetGeometry>()?.Preset?.Value;
        var outline = properties?.GetFirstChild<A.Outline>();
        if (properties is null || transform?.Offset?.X?.Value is null || transform.Offset.Y?.Value is null ||
            transform.Extents?.Cx?.Value is null or < 0 || transform.Extents.Cy?.Value is null or < 0 ||
            geometry is null || (!geometry.Equals(A.ShapeTypeValues.Line) && !geometry.Equals(A.ShapeTypeValues.BentConnector3)) ||
            outline is null || !SimpleFill(outline) ||
            outline.ChildElements.Any(child => child is not A.NoFill and not A.SolidFill and not A.HeadEnd and not A.TailEnd) ||
            properties.ChildElements.Any(child => child is not A.Transform2D and not A.PresetGeometry and not A.Outline)) return false;
        var head = outline.GetFirstChild<A.HeadEnd>();
        var tail = outline.GetFirstChild<A.TailEnd>();
        if (!TryArrow(head?.Type?.Value, out var startArrow) || !TryArrow(tail?.Type?.Value, out var endArrow)) return false;
        var nonVisual = source.NonVisualConnectionShapeProperties?.NonVisualConnectorShapeDrawingProperties;
        if (!TryConnectionTarget(nonVisual?.StartConnection, elementIdsByNativeId, out var startTargetId) ||
            !TryConnectionTarget(nonVisual?.EndConnection, elementIdsByNativeId, out var endTargetId)) return false;
        var left = transform.Offset.X.Value;
        var top = transform.Offset.Y.Value;
        var width = transform.Extents.Cx.Value;
        var height = transform.Extents.Cy.Value;
        var flipH = transform.HorizontalFlip?.Value == true;
        var flipV = transform.VerticalFlip?.Value == true;
        connector = new PresentationConnector
        {
            ConnectorType = geometry.Equals(A.ShapeTypeValues.BentConnector3) ? "elbow" : "straight",
            StartXEmu = flipH ? left + width : left,
            StartYEmu = flipV ? top + height : top,
            EndXEmu = flipH ? left : left + width,
            EndYEmu = flipV ? top : top + height,
            LineRgb = PptxColor.SolidRgb(outline.GetFirstChild<A.SolidFill>()),
            LineWidthEmu = outline.Width?.Value ?? 0,
            StartArrow = startArrow,
            EndArrow = endArrow,
            StartTargetId = startTargetId,
            EndTargetId = endTargetId,
        };
        return true;
    }

    private static P.ConnectionShape BuildConnector(PresentationElement source, uint nativeId, IReadOnlyDictionary<string, uint> nativeIdsByElementId)
    {
        ValidateConnector(source.Connector, source.Id, source.Name, nativeIdsByElementId);
        var semantic = source.Connector;
        var drawingProperties = new P.NonVisualConnectorShapeDrawingProperties();
        ApplyConnectionTargets(drawingProperties, semantic, nativeIdsByElementId);
        var properties = new P.ShapeProperties();
        properties.Append(ConnectorTransform(semantic));
        properties.Append(new A.PresetGeometry(new A.AdjustValueList()) { Preset = semantic.ConnectorType == "elbow" ? A.ShapeTypeValues.BentConnector3 : A.ShapeTypeValues.Line });
        properties.Append(ConnectorOutline(semantic));
        return new P.ConnectionShape(
            new P.NonVisualConnectionShapeProperties(
                new P.NonVisualDrawingProperties { Id = nativeId, Name = source.Name },
                drawingProperties,
                new P.ApplicationNonVisualDrawingProperties()),
            properties);
    }

    private static void ApplyConnector(P.ConnectionShape source, PresentationElement requested, IReadOnlyDictionary<string, uint> nativeIdsByElementId)
    {
        ValidateConnector(requested.Connector, requested.Id, requested.Name, nativeIdsByElementId);
        source.NonVisualConnectionShapeProperties!.NonVisualDrawingProperties!.Name = requested.Name;
        var drawingProperties = source.NonVisualConnectionShapeProperties.NonVisualConnectorShapeDrawingProperties ??= new P.NonVisualConnectorShapeDrawingProperties();
        ApplyConnectionTargets(drawingProperties, requested.Connector, nativeIdsByElementId);
        var properties = source.ShapeProperties ??= new P.ShapeProperties();
        properties.RemoveAllChildren<A.Transform2D>();
        properties.PrependChild(ConnectorTransform(requested.Connector));
        var geometry = properties.GetFirstChild<A.PresetGeometry>() ?? properties.InsertAfter(new A.PresetGeometry(new A.AdjustValueList()), properties.Transform2D);
        geometry.Preset = requested.Connector.ConnectorType == "elbow" ? A.ShapeTypeValues.BentConnector3 : A.ShapeTypeValues.Line;
        properties.GetFirstChild<A.Outline>()?.Remove();
        properties.Append(ConnectorOutline(requested.Connector));
    }

    private static bool ApplyGroup(
        P.GroupShape source,
        PresentationElement original,
        PresentationElement requested,
        PptxPartContext slideContext,
        IReadOnlyDictionary<uint, string> elementIdsByNativeId,
        IReadOnlyDictionary<string, uint> nativeIdsByElementId,
        ISet<string> changedParts,
        IDictionary<string, string> replacedOpaquePartHashes,
        int slideIndex,
        string location)
    {
        if (original.ContentCase != PresentationElement.ContentOneofCase.Group || requested.ContentCase != PresentationElement.ContentOneofCase.Group)
            throw new CodecException("presentation_group_content_changed", $"Presentation slide {slideIndex + 1} {location} changed its group content type.", PartPath(slideContext.Owner));
        var sourceChildren = GroupElements(source);
        if (sourceChildren.Length != original.Group.Children.Count || sourceChildren.Length != requested.Group.Children.Count)
            throw new CodecException("presentation_group_topology_changed", $"Presentation slide {slideIndex + 1} {location} changed its fixed group topology.", PartPath(slideContext.Owner));

        var changed = false;
        if (requested.Name != original.Name ||
            requested.Group.LeftEmu != original.Group.LeftEmu || requested.Group.TopEmu != original.Group.TopEmu ||
            requested.Group.WidthEmu != original.Group.WidthEmu || requested.Group.HeightEmu != original.Group.HeightEmu ||
            requested.Group.ChildLeftEmu != original.Group.ChildLeftEmu || requested.Group.ChildTopEmu != original.Group.ChildTopEmu ||
            requested.Group.ChildWidthEmu != original.Group.ChildWidthEmu || requested.Group.ChildHeightEmu != original.Group.ChildHeightEmu)
        {
            source.NonVisualGroupShapeProperties!.NonVisualDrawingProperties!.Name = requested.Name;
            var transform = source.GroupShapeProperties!.GetFirstChild<A.TransformGroup>()!;
            transform.Offset!.X = requested.Group.LeftEmu;
            transform.Offset.Y = requested.Group.TopEmu;
            transform.Extents!.Cx = requested.Group.WidthEmu;
            transform.Extents.Cy = requested.Group.HeightEmu;
            transform.ChildOffset!.X = requested.Group.ChildLeftEmu;
            transform.ChildOffset.Y = requested.Group.ChildTopEmu;
            transform.ChildExtents!.Cx = requested.Group.ChildWidthEmu;
            transform.ChildExtents.Cy = requested.Group.ChildHeightEmu;
            changed = true;
        }

        for (var index = 0; index < sourceChildren.Length; index++)
        {
            var sourceChild = sourceChildren[index];
            var originalChild = original.Group.Children[index];
            var requestedChild = requested.Group.Children[index];
            var binding = requestedChild.Source ?? throw new CodecException(
                "missing_presentation_element_binding",
                $"Presentation slide {slideIndex + 1} {location} child {index + 1} is missing its source binding.",
                PartPath(slideContext.Owner));
            if (requestedChild.Id != originalChild.Id || binding.ShapeTreeIndex != index ||
                !binding.ElementSha256.Equals(HashElement(sourceChild), StringComparison.OrdinalIgnoreCase) ||
                binding.Editable != originalChild.Source?.Editable ||
                binding.TextEditable != originalChild.Source?.TextEditable ||
                !binding.SemanticSha256.Equals(originalChild.Source?.SemanticSha256 ?? string.Empty, StringComparison.OrdinalIgnoreCase) ||
                !SemanticHash(originalChild).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_element_binding_mismatch",
                    $"Presentation slide {slideIndex + 1} {location} child {index + 1} does not match its owner-local source binding.",
                    PartPath(slideContext.Owner));
            if (SemanticHash(requestedChild).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase)) continue;
            if (!binding.Editable)
                throw new CodecException("unsupported_presentation_edit", $"Presentation slide {slideIndex + 1} {location} child {index + 1} is read-only.", PartPath(slideContext.Owner));
            ApplyGroupChild(sourceChild, originalChild, requestedChild, slideContext, elementIdsByNativeId, nativeIdsByElementId, changedParts, replacedOpaquePartHashes, slideIndex, $"{location} child {index + 1}");
            changed = true;
        }
        return changed;
    }

    private static void ApplyGroupChild(
        OpenXmlElement source,
        PresentationElement original,
        PresentationElement requested,
        PptxPartContext slideContext,
        IReadOnlyDictionary<uint, string> elementIdsByNativeId,
        IReadOnlyDictionary<string, uint> nativeIdsByElementId,
        ISet<string> changedParts,
        IDictionary<string, string> replacedOpaquePartHashes,
        int slideIndex,
        string location)
    {
        if (source is P.Shape shape && requested.ContentCase == PresentationElement.ContentOneofCase.Shape && IsSimpleShape(shape))
            ApplyShape(shape, requested, slideContext);
        else if (source is P.Picture picture && requested.ContentCase == PresentationElement.ContentOneofCase.Image && PptxPictureCodec.TryRead(picture, slideContext, out _))
            PptxPictureCodec.Apply(picture, requested, slideContext);
        else if (source is P.GraphicFrame table && requested.ContentCase == PresentationElement.ContentOneofCase.Table && PptxTableCodec.TryRead(table, out _))
            PptxTableCodec.Apply(table, requested);
        else if (source is P.ConnectionShape connector && requested.ContentCase == PresentationElement.ContentOneofCase.Connector && TryReadConnector(connector, elementIdsByNativeId, out _))
            ApplyConnector(connector, requested, nativeIdsByElementId);
        else if (source is P.GraphicFrame chart && requested.ContentCase == PresentationElement.ContentOneofCase.Chart && PptxChartCodec.TryRead(chart, slideContext, out _, out var chartEditable) && chartEditable)
        {
            var replacement = PptxChartCodec.Apply(chart, requested, slideContext);
            changedParts.Add(replacement.PartPath);
            replacedOpaquePartHashes.Add(replacement.PartPath, replacement.Sha256);
        }
        else if (source is P.GroupShape group && requested.ContentCase == PresentationElement.ContentOneofCase.Group && original.ContentCase == PresentationElement.ContentOneofCase.Group && TryReadGroup(group, original.Id, slideContext, elementIdsByNativeId, out _))
            _ = ApplyGroup(group, original, requested, slideContext, elementIdsByNativeId, nativeIdsByElementId, changedParts, replacedOpaquePartHashes, slideIndex, location);
        else
            throw new CodecException("unsupported_presentation_edit", $"Presentation slide {slideIndex + 1} {location} changed outside the bounded group-child profile.", PartPath(slideContext.Owner));
    }

    private static A.Transform2D ConnectorTransform(PresentationConnector source)
    {
        var left = Math.Min(source.StartXEmu, source.EndXEmu);
        var top = Math.Min(source.StartYEmu, source.EndYEmu);
        return new A.Transform2D(
            new A.Offset { X = left, Y = top },
            new A.Extents { Cx = Math.Abs(source.EndXEmu - source.StartXEmu), Cy = Math.Abs(source.EndYEmu - source.StartYEmu) })
        {
            HorizontalFlip = source.EndXEmu < source.StartXEmu,
            VerticalFlip = source.EndYEmu < source.StartYEmu,
        };
    }

    private static A.Outline ConnectorOutline(PresentationConnector source)
    {
        var outline = new A.Outline { Width = checked((int)source.LineWidthEmu) };
        outline.Append(string.IsNullOrWhiteSpace(source.LineRgb) ? new A.NoFill() : new A.SolidFill(new A.RgbColorModelHex { Val = PptxColor.Normalize(source.LineRgb) }));
        if (source.StartArrow.Length > 0) outline.Append(new A.HeadEnd { Type = A.LineEndValues.Triangle });
        if (source.EndArrow.Length > 0) outline.Append(new A.TailEnd { Type = A.LineEndValues.Triangle });
        return outline;
    }

    private static void ValidateConnector(PresentationConnector? source, string elementId, string name, IReadOnlyDictionary<string, uint>? nativeIdsByElementId = null)
    {
        if (source is null) throw new CodecException("invalid_presentation_connector", $"Presentation connector {elementId} payload is missing.");
        if (name.Length > 1_024) throw new CodecException("invalid_presentation_connector", $"Presentation connector {elementId} name exceeds 1024 characters.");
        if (source.ConnectorType is not ("straight" or "elbow")) throw new CodecException("unsupported_presentation_connector", $"Presentation connector {elementId} uses unsupported type {source.ConnectorType}.");
        if (source.StartXEmu < 0 || source.StartYEmu < 0 || source.EndXEmu < 0 || source.EndYEmu < 0 ||
            source.LineWidthEmu < 0 || source.LineWidthEmu > int.MaxValue)
            throw new CodecException("invalid_presentation_connector", $"Presentation connector {elementId} has invalid endpoints or line width.");
        if (!string.IsNullOrWhiteSpace(source.LineRgb)) PptxColor.Normalize(source.LineRgb);
        if (source.StartArrow is not ("" or "triangle") || source.EndArrow is not ("" or "triangle"))
            throw new CodecException("unsupported_presentation_connector", $"Presentation connector {elementId} uses an unsupported arrowhead.");
        if (nativeIdsByElementId is not null)
        {
            if (source.StartTargetId.Length > 0 && !nativeIdsByElementId.ContainsKey(source.StartTargetId)) throw new CodecException("invalid_presentation_connector", $"Presentation connector {elementId} references missing start target {source.StartTargetId}.");
            if (source.EndTargetId.Length > 0 && !nativeIdsByElementId.ContainsKey(source.EndTargetId)) throw new CodecException("invalid_presentation_connector", $"Presentation connector {elementId} references missing end target {source.EndTargetId}.");
        }
    }

    private static bool TryArrow(A.LineEndValues? source, out string arrow)
    {
        arrow = string.Empty;
        if (source is null || source.Value == A.LineEndValues.None) return true;
        if (source.Value != A.LineEndValues.Triangle) return false;
        arrow = "triangle";
        return true;
    }

    private static bool TryConnectionTarget(A.StartConnection? source, IReadOnlyDictionary<uint, string>? ids, out string targetId) =>
        TryConnectionTarget(source?.Id?.Value, ids, out targetId);

    private static bool TryConnectionTarget(A.EndConnection? source, IReadOnlyDictionary<uint, string>? ids, out string targetId) =>
        TryConnectionTarget(source?.Id?.Value, ids, out targetId);

    private static bool TryConnectionTarget(uint? nativeId, IReadOnlyDictionary<uint, string>? ids, out string targetId)
    {
        targetId = string.Empty;
        if (nativeId is null) return true;
        return ids is not null && ids.TryGetValue(nativeId.Value, out targetId!);
    }

    private static void ApplyConnectionTargets(P.NonVisualConnectorShapeDrawingProperties properties, PresentationConnector source, IReadOnlyDictionary<string, uint> nativeIdsByElementId)
    {
        properties.RemoveAllChildren<A.StartConnection>();
        properties.RemoveAllChildren<A.EndConnection>();
        if (source.StartTargetId.Length > 0) properties.Append(new A.StartConnection { Id = nativeIdsByElementId[source.StartTargetId], Index = 0U });
        if (source.EndTargetId.Length > 0) properties.Append(new A.EndConnection { Id = nativeIdsByElementId[source.EndTargetId], Index = 0U });
    }

    private static P.ShapeTree BasicShapeTree() => new(
        new P.NonVisualGroupShapeProperties(
            new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
            new P.NonVisualGroupShapeDrawingProperties(),
            new P.ApplicationNonVisualDrawingProperties()),
        new P.GroupShapeProperties(new A.TransformGroup(
            new A.Offset { X = 0L, Y = 0L },
            new A.Extents { Cx = 0L, Cy = 0L },
            new A.ChildOffset { X = 0L, Y = 0L },
            new A.ChildExtents { Cx = 0L, Cy = 0L })));

    private static P.ColorMap BasicColorMap() => new()
    {
        Background1 = A.ColorSchemeIndexValues.Light1,
        Text1 = A.ColorSchemeIndexValues.Dark1,
        Background2 = A.ColorSchemeIndexValues.Light2,
        Text2 = A.ColorSchemeIndexValues.Dark2,
        Accent1 = A.ColorSchemeIndexValues.Accent1,
        Accent2 = A.ColorSchemeIndexValues.Accent2,
        Accent3 = A.ColorSchemeIndexValues.Accent3,
        Accent4 = A.ColorSchemeIndexValues.Accent4,
        Accent5 = A.ColorSchemeIndexValues.Accent5,
        Accent6 = A.ColorSchemeIndexValues.Accent6,
        Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
        FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink,
    };

    private static A.Theme BasicTheme() => new(
        new A.ThemeElements(
            new A.ColorScheme(
                new A.Dark1Color(new A.SystemColor { Val = A.SystemColorValues.WindowText, LastColor = "000000" }),
                new A.Light1Color(new A.SystemColor { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }),
                new A.Dark2Color(new A.RgbColorModelHex { Val = "1F497D" }),
                new A.Light2Color(new A.RgbColorModelHex { Val = "EEECE1" }),
                new A.Accent1Color(new A.RgbColorModelHex { Val = "4F81BD" }),
                new A.Accent2Color(new A.RgbColorModelHex { Val = "C0504D" }),
                new A.Accent3Color(new A.RgbColorModelHex { Val = "9BBB59" }),
                new A.Accent4Color(new A.RgbColorModelHex { Val = "8064A2" }),
                new A.Accent5Color(new A.RgbColorModelHex { Val = "4BACC6" }),
                new A.Accent6Color(new A.RgbColorModelHex { Val = "F79646" }),
                new A.Hyperlink(new A.RgbColorModelHex { Val = "0000FF" }),
                new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = "800080" })) { Name = "Office" },
            new A.FontScheme(
                new A.MajorFont(new A.LatinFont { Typeface = "Arial" }, new A.EastAsianFont { Typeface = string.Empty }, new A.ComplexScriptFont { Typeface = string.Empty }),
                new A.MinorFont(new A.LatinFont { Typeface = "Arial" }, new A.EastAsianFont { Typeface = string.Empty }, new A.ComplexScriptFont { Typeface = string.Empty })) { Name = "Office" },
            new A.FormatScheme(
                new A.FillStyleList(
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })),
                new A.LineStyleList(
                    BasicThemeOutline(9_525),
                    BasicThemeOutline(25_400),
                    BasicThemeOutline(38_100)),
                new A.EffectStyleList(
                    new A.EffectStyle(new A.EffectList()),
                    new A.EffectStyle(new A.EffectList()),
                    new A.EffectStyle(new A.EffectList())),
                new A.BackgroundFillStyleList(
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }))) { Name = "Office" }))
    { Name = "Office Clean Room" };

    private static A.Outline BasicThemeOutline(int width) => new(
        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
        new A.PresetDash { Val = A.PresetLineDashValues.Solid })
    { Width = width, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

    private static OpenXmlElement[] ShapeElements(P.ShapeTree shapeTree) =>
        shapeTree.ChildElements.Where(child => child is not P.NonVisualGroupShapeProperties and not P.GroupShapeProperties).ToArray();

    private static OpenXmlElement[] GroupElements(P.GroupShape group) =>
        group.ChildElements.Where(child => child is not P.NonVisualGroupShapeProperties and not P.GroupShapeProperties).ToArray();

    private static IReadOnlyDictionary<uint, string> NativeElementIds(IReadOnlyList<OpenXmlElement> elements, string ownerId)
    {
        var output = new Dictionary<uint, string>();
        CollectNativeElementIds(elements, ownerId, output);
        return output;
    }

    private static void CollectNativeElementIds(IReadOnlyList<OpenXmlElement> elements, string ownerId, IDictionary<uint, string> output)
    {
        for (var index = 0; index < elements.Count; index++)
        {
            var elementId = $"{ownerId}/element/{index + 1}";
            var nativeId = elements[index].Descendants<P.NonVisualDrawingProperties>().FirstOrDefault()?.Id?.Value;
            if (nativeId is not null) output[nativeId.Value] = elementId;
            if (elements[index] is P.GroupShape group)
                CollectNativeElementIds(GroupElements(group), elementId, output);
        }
    }

    private static (long Left, long Top, long Width, long Height) ReadFrame(OpenXmlElement element)
    {
        if (element is P.GraphicFrame graphicFrame && graphicFrame.Transform?.Offset is { } graphicOffset && graphicFrame.Transform.Extents is { } graphicExtents)
            return (graphicOffset.X?.Value ?? 0, graphicOffset.Y?.Value ?? 0, graphicExtents.Cx?.Value ?? 0, graphicExtents.Cy?.Value ?? 0);
        if (element is P.GroupShape group && group.GetFirstChild<P.GroupShapeProperties>()?.GetFirstChild<A.TransformGroup>() is { Offset: { } groupOffset, Extents: { } groupExtents })
            return (groupOffset.X?.Value ?? 0, groupOffset.Y?.Value ?? 0, groupExtents.Cx?.Value ?? 0, groupExtents.Cy?.Value ?? 0);
        var transform = element.Descendants<A.Transform2D>().FirstOrDefault();
        if (transform?.Offset is not null && transform.Extents is not null)
            return (transform.Offset.X?.Value ?? 0, transform.Offset.Y?.Value ?? 0, transform.Extents.Cx?.Value ?? 0, transform.Extents.Cy?.Value ?? 0);
        var offset = element.Descendants<A.Offset>().FirstOrDefault();
        var extents = element.Descendants<A.Extents>().FirstOrDefault();
        return (offset?.X?.Value ?? 0, offset?.Y?.Value ?? 0, extents?.Cx?.Value ?? 0, extents?.Cy?.Value ?? 0);
    }

    private static string Geometry(P.Shape shape)
    {
        if (shape.NonVisualShapeProperties?.NonVisualShapeDrawingProperties?.TextBox?.Value == true) return "textbox";
        if (shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>() is not null) return "custom";
        var value = shape.ShapeProperties?.GetFirstChild<A.PresetGeometry>()?.Preset?.Value;
        if (value is null) return "rect";
        return value.Equals(A.ShapeTypeValues.Ellipse) ? "ellipse" :
            value.Equals(A.ShapeTypeValues.RoundRectangle) ? "roundRect" :
            value.Equals(A.ShapeTypeValues.Rectangle) ? "rect" : value.ToString() ?? "rect";
    }

    private static string ElementName(OpenXmlElement element, int index) =>
        element.Descendants<P.NonVisualDrawingProperties>().FirstOrDefault()?.Name?.Value ?? $"{element.LocalName} {index + 1}";

    private static string DescendantText(OpenXmlElement? element) =>
        element is null ? string.Empty : string.Concat(element.Descendants<A.Text>().Select(text => text.Text));

    private static string SemanticHash(PresentationElement element)
    {
        var semantic = element.Clone();
        var placementEditable = semantic.ContentCase == PresentationElement.ContentOneofCase.Opaque && semantic.Source?.Editable == true;
        ClearElementIdentity(semantic);
        if (semantic.ContentCase == PresentationElement.ContentOneofCase.Shape) PptxTextCodec.NormalizeSemantics(semantic.Shape);
        else if (placementEditable) semantic.Opaque.RawXml = string.Empty;
        return Hash(semantic.ToByteArray());
    }

    private static void ClearElementIdentity(PresentationElement element)
    {
        element.Id = string.Empty;
        element.Source = null;
        if (element.ContentCase != PresentationElement.ContentOneofCase.Group) return;
        foreach (var child in element.Group.Children) ClearElementIdentity(child);
    }

    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static string RelationshipPartPath(OpenXmlPart part)
    {
        var path = PartPath(part);
        var separator = path.LastIndexOf('/');
        var directory = separator < 0 ? string.Empty : path[..separator];
        var fileName = separator < 0 ? path : path[(separator + 1)..];
        return directory.Length == 0 ? $"_rels/{fileName}.rels" : $"{directory}/_rels/{fileName}.rels";
    }
    private static byte[] PartBytes(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }
    private static void CopyPartBytes(OpenXmlPart source, OpenXmlPart target)
    {
        using var input = source.GetStream(FileMode.Open, FileAccess.Read);
        using var output = target.GetStream(FileMode.Create, FileAccess.Write);
        input.CopyTo(output);
    }
    private static string HashElement(OpenXmlElement element) => Hash(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static PptxAssetCatalog ValidateEnvelope(ArtifactEnvelope envelope, EffectiveCodecLimits limits)
    {
        if (envelope.ProtocolVersion != CodecProtocol.ProtocolVersion)
            throw new CodecException("unsupported_artifact_version", $"Artifact protocol version {envelope.ProtocolVersion} is unsupported.");
        if (envelope.Family != ArtifactFamily.Presentation || envelope.PayloadCase != ArtifactEnvelope.PayloadOneofCase.Presentation)
            throw new CodecException("invalid_presentation_artifact", "Artifact envelope does not contain a presentation payload.");
        if (envelope.Presentation.Slides.Count == 0)
            throw new CodecException("missing_slides", "Presentation must contain at least one slide.");
        if ((uint)envelope.Presentation.Slides.Count > limits.MaxSheets)
            throw new CodecException("slide_budget_exceeded", $"Presentation has {envelope.Presentation.Slides.Count} slides and exceeds max_sheets ({limits.MaxSheets}).");
        if (envelope.Presentation.SlideWidthEmu < 0 || envelope.Presentation.SlideHeightEmu < 0 || envelope.Presentation.SlideWidthEmu > int.MaxValue || envelope.Presentation.SlideHeightEmu > int.MaxValue)
            throw new CodecException("invalid_slide_size", "Presentation slide dimensions must fit the PresentationML signed 32-bit EMU range.");
        var assetCatalog = new PptxAssetCatalog(envelope.Assets, limits);
        var hasSourcePackage = envelope.OpaqueOpc?.SourcePackage is { Data.IsEmpty: false };
        PptxViewPropertiesCodec.Validate(envelope.Presentation.ViewProperties, hasSourcePackage);

        if (envelope.Presentation.Masters.Count > 64)
            throw new CodecException("presentation_master_budget_exceeded", "Presentation cannot contain more than 64 slide masters.");
        if ((uint)envelope.Presentation.Layouts.Count > limits.MaxSheets)
            throw new CodecException("presentation_layout_budget_exceeded", $"Presentation has {envelope.Presentation.Layouts.Count} layouts and exceeds max_sheets ({limits.MaxSheets}).");
        ulong items = 0;
        var masterIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var master in envelope.Presentation.Masters)
        {
            if (string.IsNullOrWhiteSpace(master.Id) || !masterIds.Add(master.Id))
                throw new CodecException("invalid_presentation_master", "Presentation master IDs must be non-empty and unique.");
            if (master.Name.Length > 1_024)
                throw new CodecException("invalid_presentation_master", $"Presentation master {master.Id} name exceeds 1024 characters.");
            PptxMasterTextStylesCodec.Validate(master.TextStyles);
            PptxBackgroundCodec.Validate(master.Background);
            ValidatePlaceholders(master.Id, master.Placeholders, assetCatalog, limits, ref items);
            foreach (var paragraph in MasterStyleParagraphs(master.TextStyles))
                if (paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.PictureBullet &&
                    paragraph.PictureBullet.SourceCase == PresentationPictureBullet.SourceOneofCase.AssetId)
                    _ = assetCatalog.Get(paragraph.PictureBullet.AssetId);
        }
        var layoutIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var layout in envelope.Presentation.Layouts)
        {
            if (string.IsNullOrWhiteSpace(layout.Id) || !layoutIds.Add(layout.Id))
                throw new CodecException("invalid_presentation_layout", "Presentation layout IDs must be non-empty and unique.");
            if (!masterIds.Contains(layout.MasterId))
                throw new CodecException("invalid_presentation_layout", $"Presentation layout {layout.Id} references missing master {layout.MasterId}.");
            if (layout.Name.Length > 1_024 || layout.Type.Length > 128)
                throw new CodecException("invalid_presentation_layout", $"Presentation layout {layout.Id} has invalid name or type metadata.");
            PptxBackgroundCodec.Validate(layout.Background);
            ValidatePlaceholders(layout.Id, layout.Placeholders, assetCatalog, limits, ref items);
        }

        var mastersById = envelope.Presentation.Masters.ToDictionary(master => master.Id, StringComparer.Ordinal);
        var layoutsById = envelope.Presentation.Layouts.ToDictionary(layout => layout.Id, StringComparer.Ordinal);
        if (!hasSourcePackage)
        {
            foreach (var master in envelope.Presentation.Masters)
                foreach (var placeholder in master.Placeholders)
                    ValidateSourceFreeTextPlaceholder(placeholder, master.Id);
            foreach (var layout in envelope.Presentation.Layouts)
            {
                _ = SourceFreeLayoutType(layout);
                foreach (var placeholder in layout.Placeholders)
                    ValidateSourceFreeTextPlaceholder(placeholder, layout.Id);
            }
        }

        for (var slideIndex = 0; slideIndex < envelope.Presentation.Slides.Count; slideIndex++)
        {
            var slide = envelope.Presentation.Slides[slideIndex];
            PptxSpeakerNotesCodec.Validate(slide.SpeakerNotes);
            PptxBackgroundCodec.Validate(slide.Background);
            PptxLegacyCommentsCodec.Validate(slide, slideIndex);
            PptxModernCommentsCodec.Validate(slide, slideIndex, hasSourcePackage);
            if (!string.IsNullOrWhiteSpace(slide.LayoutId) && !layoutIds.Contains(slide.LayoutId))
                throw new CodecException("invalid_presentation_layout", $"Presentation slide {slide.Id} references missing layout {slide.LayoutId}.");
            if (!hasSourcePackage)
            {
                var placeholderShapes = slide.Elements.Where(element =>
                    element.ContentCase == PresentationElement.ContentOneofCase.Shape && element.Shape.Placeholder is not null).ToArray();
                if (placeholderShapes.Length > 0)
                {
                    if (string.IsNullOrWhiteSpace(slide.LayoutId) || !layoutsById.TryGetValue(slide.LayoutId, out var layout))
                        throw new CodecException("invalid_presentation_layout", $"Source-free presentation slide {slide.Id} has placeholders but no explicit layout binding.");
                    var master = mastersById[layout.MasterId];
                    foreach (var element in placeholderShapes)
                    {
                        var placeholder = element.Shape.Placeholder;
                        if (placeholder.InheritsGeometry || element.Shape.DirectFrame is null)
                            throw new CodecException("invalid_presentation_placeholder", $"Source-free presentation slide placeholder {element.Id} must use a direct frame.");
                        ValidateSourceFreeTextPlaceholderIdentity(placeholder, element.Id);
                        if (!master.Placeholders.Concat(layout.Placeholders).Any(candidate =>
                                candidate.Type.Equals(placeholder.Type, StringComparison.Ordinal) && candidate.Index == placeholder.Index))
                            throw new CodecException("presentation_placeholder_binding_mismatch", $"Source-free presentation slide placeholder {element.Id} has no matching master/layout placeholder.");
                    }
                }
            }
            foreach (var element in slide.Elements)
                ValidatePresentationElement(element, hasSourcePackage, assetCatalog, limits, ref items, 0);
        }
        return assetCatalog;
    }

    private static void ValidatePresentationElement(
        PresentationElement element,
        bool hasSourcePackage,
        PptxAssetCatalog assetCatalog,
        EffectiveCodecLimits limits,
        ref ulong items,
        int depth)
    {
        items++;
        if (items > limits.MaxCells)
            throw new CodecException("presentation_item_budget_exceeded", $"Presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).");
        if (depth > 16)
            throw new CodecException("presentation_group_depth_exceeded", "Presentation groups cannot be nested more than 16 levels.");
        if (string.IsNullOrWhiteSpace(element.Id) || element.Id.Length > 1_024 || element.Name.Length > 1_024)
            throw new CodecException("invalid_presentation_element", "Presentation element IDs and names must be bounded non-empty metadata.");

        // The source-preserving path re-reads this exact bound element before
        // applying any request, verifies its native and semantic hashes, and
        // rejects every semantic change when Editable is false. Treat that
        // contract as one opaque branch here rather than re-implementing a
        // partial validator for every DrawingML feature we deliberately do
        // not model. The envelope-wide item/depth/id bounds above still apply.
        if (hasSourcePackage && element.Source?.Editable == false) return;

        if (element.ContentCase == PresentationElement.ContentOneofCase.Shape)
        {
            if (element.Shape.HasUseBackgroundFill && !hasSourcePackage)
                throw new CodecException("unsupported_presentation_features", $"Presentation shape {element.Id} cannot author useBgFill without a validated source package.");
            if (element.Shape.Placeholder is not null && !hasSourcePackage &&
                (element.Shape.Placeholder.InheritsGeometry || element.Shape.DirectFrame is null))
                throw new CodecException("invalid_presentation_placeholder", $"Source-free presentation slide placeholder {element.Id} must use a direct frame.");
            if (element.Shape.Placeholder is not null && element.Shape.Transform is not null)
                throw new CodecException("invalid_presentation_transform", $"Presentation placeholder shape {element.Id} cannot carry an ordinary shape transform.");
            var inheritedPlaceholderGeometry = element.Shape.Placeholder?.InheritsGeometry == true &&
                element.Shape.DirectFrame is null && element.Source?.Editable == false;
            if ((!inheritedPlaceholderGeometry && (element.Shape.LeftEmu < 0 || element.Shape.TopEmu < 0 || element.Shape.WidthEmu <= 0 || element.Shape.HeightEmu <= 0)) ||
                element.Shape.LineWidthEmu < 0 || element.Shape.LineWidthEmu > int.MaxValue)
                throw new CodecException("invalid_presentation_frame", $"Presentation shape {element.Id} has an invalid frame.");
            if (element.Shape.DirectFrame is not null)
            {
                if (element.Shape.Placeholder is null || element.Shape.Placeholder.InheritsGeometry)
                    throw new CodecException("invalid_presentation_placeholder", $"Presentation shape {element.Id} has inconsistent direct placeholder geometry.");
                PptxPlaceholderCodec.ValidateDirectFrame(element.Shape.DirectFrame, element.Id);
            }
            if (element.Shape.Geometry is not ("rect" or "ellipse" or "roundRect" or "textbox" or "custom"))
                throw new CodecException("unsupported_presentation_geometry", $"Presentation shape {element.Id} uses unsupported geometry {element.Shape.Geometry}.");
            PptxCustomGeometryCodec.Validate(element.Shape, element.Id);
            if (!string.IsNullOrWhiteSpace(element.Shape.FillRgb)) PptxColor.Normalize(element.Shape.FillRgb);
            if (!string.IsNullOrWhiteSpace(element.Shape.LineRgb)) PptxColor.Normalize(element.Shape.LineRgb);
            PptxShapeTransformCodec.Validate(element.Shape.Transform, element.Id);
            ValidateShadow(element.Shape.Shadow, element.Id);
            PptxTextCodec.Validate(element.Shape);
            foreach (var paragraph in element.Shape.TextBody?.Paragraphs ?? [])
                if (paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.PictureBullet &&
                    paragraph.PictureBullet.SourceCase == PresentationPictureBullet.SourceOneofCase.AssetId)
                    _ = assetCatalog.Get(paragraph.PictureBullet.AssetId);
        }
        else if (element.ContentCase == PresentationElement.ContentOneofCase.Image)
            PptxPictureCodec.Validate(element.Image, element.Id, assetCatalog);
        else if (element.ContentCase == PresentationElement.ContentOneofCase.Table)
        {
            PptxTableCodec.Validate(element.Table, element.Id);
            items += checked((ulong)element.Table.Rows.Sum(row => row.Cells.Count));
            if (items > limits.MaxCells)
                throw new CodecException("presentation_item_budget_exceeded", $"Presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).");
        }
        else if (element.ContentCase == PresentationElement.ContentOneofCase.Connector)
            ValidateConnector(element.Connector, element.Id, element.Name);
        else if (element.ContentCase == PresentationElement.ContentOneofCase.Chart)
        {
            PptxChartCodec.Validate(element.Chart, element.Id, element.Name);
            items += checked((ulong)(
                element.Chart.Series.Sum(series => series.Values.Count) +
                element.Chart.ComboSeries.Sum(entry => entry.Series?.Values.Count ?? 0)));
            if (items > limits.MaxCells)
                throw new CodecException("presentation_item_budget_exceeded", $"Presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).");
        }
        else if (element.ContentCase == PresentationElement.ContentOneofCase.Group)
        {
            var group = element.Group;
            if (group.LeftEmu < 0 || group.TopEmu < 0 || group.WidthEmu <= 0 || group.HeightEmu <= 0 ||
                group.ChildWidthEmu <= 0 || group.ChildHeightEmu <= 0 || group.Children.Count == 0)
                throw new CodecException("invalid_presentation_group", $"Presentation group {element.Id} requires positive outer/child extents and at least one child.");
            var childIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (var child in group.Children)
            {
                if (!childIds.Add(child.Id))
                    throw new CodecException("invalid_presentation_group", $"Presentation group {element.Id} contains duplicate child ID {child.Id}.");
                if (child.ContentCase == PresentationElement.ContentOneofCase.Opaque)
                    throw new CodecException("unsupported_presentation_features", $"Presentation group {element.Id} contains a source-free or semantically edited opaque child.");
                ValidatePresentationElement(child, hasSourcePackage, assetCatalog, limits, ref items, depth + 1);
            }
        }
        else if (element.ContentCase != PresentationElement.ContentOneofCase.Opaque)
            throw new CodecException("missing_presentation_element_content", $"Presentation element {element.Id} has no content.");
        else if (element.Source?.Editable == true)
        {
            if (element.Opaque.LeftEmu < 0 || element.Opaque.TopEmu < 0 || element.Opaque.WidthEmu <= 0 || element.Opaque.HeightEmu <= 0)
                throw new CodecException("invalid_presentation_frame", $"Presentation native object {element.Id} has an invalid frame.");
        }
    }

    private static void ValidatePlaceholders(
        string ownerId,
        IList<PresentationPlaceholder> placeholders,
        PptxAssetCatalog assetCatalog,
        EffectiveCodecLimits limits,
        ref ulong items)
    {
        if (placeholders.Count > 128)
            throw new CodecException("presentation_placeholder_budget_exceeded", $"Presentation owner {ownerId} exceeds the 128-placeholder budget.");
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var placeholder in placeholders)
        {
            items++;
            if (items > limits.MaxCells)
                throw new CodecException("presentation_item_budget_exceeded", $"Presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).");
            if (!ids.Add(placeholder.Id))
                throw new CodecException("invalid_presentation_placeholder", $"Presentation owner {ownerId} contains duplicate placeholder ID {placeholder.Id}.");
            PptxPlaceholderCodec.Validate(placeholder);
            foreach (var paragraph in (placeholder.TextBody?.Paragraphs ?? []).Concat(placeholder.TextBody?.ListStyles ?? []))
                if (paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.PictureBullet &&
                    paragraph.PictureBullet.SourceCase == PresentationPictureBullet.SourceOneofCase.AssetId)
                    _ = assetCatalog.Get(paragraph.PictureBullet.AssetId);
        }
    }

    private static void AssertPackagePartsUnchangedExcept(byte[] sourceBytes, byte[] outputBytes, HashSet<string> allowedPaths)
    {
        var before = PackagePartHashes(sourceBytes);
        var after = PackagePartHashes(outputBytes);
        var inventoryChanges = before.Keys.Concat(after.Keys).Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(path => !before.ContainsKey(path) || !after.ContainsKey(path))
            .Where(path => !allowedPaths.Contains(path))
            .Take(8)
            .ToArray();
        if (inventoryChanges.Length > 0)
            throw new CodecException("presentation_package_topology_changed", $"Source-preserving PPTX export changed unowned OPC part inventory: {string.Join(", ", inventoryChanges)}.");
        var changed = before.Keys.Intersect(after.Keys, StringComparer.OrdinalIgnoreCase)
            .Where(path => !before[path].Equals(after[path], StringComparison.OrdinalIgnoreCase))
            .ToArray();
        var unexpected = changed.Where(path => !allowedPaths.Contains(path)).Take(8).ToArray();
        if (unexpected.Length > 0)
            throw new CodecException("presentation_unowned_part_changed", $"Source-preserving PPTX export changed unowned package parts: {string.Join(", ", unexpected)}.");
    }

    private static void AssertOpaqueGraphMatchesWithModeledAdditions(
        OpaqueOpcGraph expected,
        OpaqueOpcGraph actual,
        IReadOnlySet<string> allowedAddedRelationshipIds,
        IReadOnlySet<string> allowedAddedPartPaths,
        IReadOnlyDictionary<string, string> allowedChangedPartHashes,
        IReadOnlySet<string> removedSourceSlidePartPaths)
    {
        var guarded = actual.Clone();
        var removed = new HashSet<string>(StringComparer.Ordinal);
        foreach (var relationship in guarded.PackageRelationships.ToArray())
        {
            var key = $"{relationship.SourcePath}\0{relationship.Id}";
            if (!allowedAddedRelationshipIds.Contains(key)) continue;
            var isExternalLink = relationship.Type.EndsWith("/hyperlink", StringComparison.Ordinal) &&
                                 relationship.TargetMode.Equals("External", StringComparison.OrdinalIgnoreCase);
            var isSlideJump = relationship.Type.EndsWith("/slide", StringComparison.Ordinal) &&
                              !relationship.TargetMode.Equals("External", StringComparison.OrdinalIgnoreCase);
            var isImage = relationship.Type.EndsWith("/image", StringComparison.Ordinal);
            // A bounded source-preserving clone owns one new internal edge:
            // its new SlidePart points at the original shared SlideLayoutPart.
            // Keep that exception typed and target-checked; an arbitrary added
            // relationship from a clone slide remains an opaque-graph breach.
            var isSlideLayout = relationship.Type.EndsWith("/slideLayout", StringComparison.Ordinal) &&
                                !relationship.TargetMode.Equals("External", StringComparison.OrdinalIgnoreCase) &&
                                IsNumberedLayoutPath(PptxNativeObjectCatalog.ResolveTarget(relationship.SourcePath, relationship.Target));
            var isNotesSlide = relationship.Type.EndsWith("/notesSlide", StringComparison.Ordinal) &&
                               !relationship.TargetMode.Equals("External", StringComparison.OrdinalIgnoreCase) &&
                               IsNumberedNotesSlidePath(PptxNativeObjectCatalog.ResolveTarget(relationship.SourcePath, relationship.Target));
            var isNotesMaster = relationship.Type.EndsWith("/notesMaster", StringComparison.Ordinal) &&
                                !relationship.TargetMode.Equals("External", StringComparison.OrdinalIgnoreCase) &&
                                IsNumberedNotesMasterPath(PptxNativeObjectCatalog.ResolveTarget(relationship.SourcePath, relationship.Target));
            var isNotesBackReference = relationship.Type.EndsWith("/slide", StringComparison.Ordinal) &&
                                       !relationship.TargetMode.Equals("External", StringComparison.OrdinalIgnoreCase) &&
                                       IsNumberedSlidePath(PptxNativeObjectCatalog.ResolveTarget(relationship.SourcePath, relationship.Target));
            var isLegacyComments = relationship.Type.EndsWith("/comments", StringComparison.Ordinal) &&
                                   !relationship.TargetMode.Equals("External", StringComparison.OrdinalIgnoreCase) &&
                                   IsNumberedCommentsPath(PptxNativeObjectCatalog.ResolveTarget(relationship.SourcePath, relationship.Target));
            var allowedFromSlide = IsNumberedSlidePath(relationship.SourcePath) && (isExternalLink || isSlideJump || isImage || isSlideLayout || isNotesSlide || isLegacyComments);
            var allowedFromMaster = IsNumberedMasterPath(relationship.SourcePath) && (isExternalLink || isSlideJump || isImage);
            var allowedFromLayout = IsNumberedLayoutPath(relationship.SourcePath) && (isExternalLink || isSlideJump || isImage);
            var allowedFromNotes = IsNumberedNotesSlidePath(relationship.SourcePath) && (isNotesMaster || isNotesBackReference);
            if (!allowedFromSlide && !allowedFromMaster && !allowedFromLayout && !allowedFromNotes)
                throw new CodecException("opaque_content_not_preserved", $"Modeled PPTX edit added unsupported relationship {relationship.Id} from {relationship.SourcePath}.");
            guarded.PackageRelationships.Remove(relationship);
            removed.Add(key);
        }
        if (!removed.SetEquals(allowedAddedRelationshipIds))
            throw new CodecException("opaque_content_not_preserved", "Modeled PPTX relationship additions do not match the relationships written to the package.");
        var removedParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var part in guarded.Parts.ToArray())
        {
            if (!allowedAddedPartPaths.Contains(part.Path)) continue;
            if (!IsAllowedAddedClonePartPath(part.Path))
                throw new CodecException("opaque_content_not_preserved", $"Modeled PPTX edit added unsupported part {part.Path}.");
            guarded.Parts.Remove(part);
            removedParts.Add(part.Path);
        }
        if (!removedParts.SetEquals(allowedAddedPartPaths))
            throw new CodecException("opaque_content_not_preserved", "Modeled PPTX clone additions do not match the parts written to the package.");
        foreach (var (path, requestedHash) in allowedChangedPartHashes)
        {
            var before = expected.Parts.SingleOrDefault(part => part.Path.Equals(path, StringComparison.OrdinalIgnoreCase));
            var after = guarded.Parts.SingleOrDefault(part => part.Path.Equals(path, StringComparison.OrdinalIgnoreCase));
            if (before is null || after is null ||
                !before.ContentType.Equals(after.ContentType, StringComparison.OrdinalIgnoreCase) ||
                !after.Sha256.Equals(requestedHash, StringComparison.OrdinalIgnoreCase) ||
                !before.Relationships.SequenceEqual(after.Relationships))
                throw new CodecException("opaque_content_not_preserved", $"Modeled PPTX OLE workbook replacement did not preserve the package contract for {path}.", path);
        }
        PackageGuards.AssertOpaqueGraphMatches(
            expected,
            guarded,
            "opaque_content_not_preserved",
            ignoreRelationship: relationship => removedSourceSlidePartPaths.Contains(relationship.SourcePath),
            ignorePart: part => allowedChangedPartHashes.ContainsKey(part.Path));
    }

    private static bool IsNumberedSlidePath(string path)
    {
        const string prefix = "ppt/slides/slide";
        const string suffix = ".xml";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsNumberedMasterPath(string path)
    {
        const string prefix = "ppt/slideMasters/slideMaster";
        const string suffix = ".xml";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsNumberedLayoutPath(string path)
    {
        const string prefix = "ppt/slideLayouts/slideLayout";
        const string suffix = ".xml";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsNumberedNotesSlidePath(string path)
    {
        const string prefix = "ppt/notesSlides/notesSlide";
        const string suffix = ".xml";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsNumberedNotesSlideRelationshipPath(string path)
    {
        const string prefix = "ppt/notesSlides/_rels/notesSlide";
        const string suffix = ".xml.rels";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsNumberedNotesMasterPath(string path)
    {
        const string prefix = "ppt/notesMasters/notesMaster";
        const string suffix = ".xml";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsNumberedCommentsPath(string path)
    {
        const string prefix = "ppt/comments/comment";
        const string suffix = ".xml";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsAllowedAddedClonePartPath(string path) =>
        path.StartsWith("ppt/media/", StringComparison.OrdinalIgnoreCase) ||
        IsNumberedNotesSlidePath(path) ||
        IsNumberedNotesSlideRelationshipPath(path) ||
        IsNumberedCommentsPath(path);

    private static void ValidatePreservedSlideElements(byte[] sourceBytes, byte[] outputBytes, PresentationArtifact requested, EffectiveCodecLimits limits)
    {
        using var sourceStream = new MemoryStream(sourceBytes, writable: false);
        using var outputStream = new MemoryStream(outputBytes, writable: false);
        using var sourcePackage = PresentationDocument.Open(sourceStream, isEditable: false);
        using var outputPackage = PresentationDocument.Open(outputStream, isEditable: false);
        var sourcePresentationPart = sourcePackage.PresentationPart ??
            throw new CodecException("missing_presentation_part", "PPTX source package has no Presentation part.", "ppt/presentation.xml");
        var outputPresentationPart = outputPackage.PresentationPart ??
            throw new CodecException("missing_presentation_part", "PPTX output package has no Presentation part.", "ppt/presentation.xml");
        var sourceTargets = BindSourcePreservingSlides(
            sourcePresentationPart,
            sourcePresentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>().ToArray() ?? [],
            requested.Slides);
        var outputSlides = OrderedSlideParts(outputPackage);
        if (sourceTargets.Length != requested.Slides.Count || outputSlides.Length != requested.Slides.Count)
            throw new CodecException("presentation_postwrite_topology_changed", "PPTX slide topology changed during source-preserving export.");
        for (var targetIndex = 0; targetIndex < sourceTargets.Length; targetIndex++)
        {
            var target = sourceTargets[targetIndex];
            var outputSlide = outputSlides[targetIndex];
            if (!target.IsClone && !PartPath(outputSlide).Equals(PartPath(target.Source.Part), StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_postwrite_topology_changed", "PPTX slide order does not match the requested source-bound order.", "ppt/presentation.xml");
            if (target.IsClone)
            {
                if (PartPath(outputSlide).Equals(PartPath(target.Source.Part), StringComparison.OrdinalIgnoreCase) ||
                    !HashElement(target.Source.Part.Slide!).Equals(HashElement(outputSlide.Slide!), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_postwrite_clone_mismatch", $"PPTX clone {targetIndex + 1} is not an independent exact source slide copy.", PartPath(outputSlide));
                ValidateClonedRunHyperlinkGraph(target.Source.Part, outputSlide, targetIndex);
                ValidateClonedSpeakerNotes(target.Source.Part, outputSlide, targetIndex);
                ValidateClonedLegacyComments(sourcePresentationPart, outputPresentationPart, target.Source.Part, outputSlide, targetIndex);
            }
        }
        var retainedTargets = sourceTargets.Where(target => !target.IsClone).ToArray();
        var sourceIdByPartPath = retainedTargets
            .Select(target => (Path: PartPath(target.Source.Part), Id: target.Target.Id))
            .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
        var outputIdByPartPath = retainedTargets
            .Select(target => (Path: PartPath(target.Source.Part), Id: target.Target.Id))
            .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
        var sourceAssets = new PptxAssetCatalog([], limits);
        var outputAssets = new PptxAssetCatalog([], limits);

        for (var slideIndex = 0; slideIndex < requested.Slides.Count; slideIndex++)
        {
            if (sourceTargets[slideIndex].IsClone) continue;
            var sourceSlide = sourceTargets[slideIndex].Source.Part;
            var outputSlide = outputSlides[slideIndex];
            var outputRoot = outputSlide.Slide ??
                throw new CodecException("missing_slide_root", $"PPTX output slide {slideIndex + 1} has no slide root.", PartPath(outputSlide));
            var outputName = outputRoot.CommonSlideData?.Name?.Value ?? $"Slide {sourceTargets[slideIndex].Source.Index + 1}";
            if (!string.Equals(outputName, requested.Slides[slideIndex].Name, StringComparison.Ordinal))
                throw new CodecException(
                    "presentation_postwrite_slide_name_mismatch",
                    $"PPTX slide {slideIndex + 1} name does not match the requested source-bound value.",
                    PartPath(outputSlide));
            var sourceContext = new PptxPartContext(sourceSlide, sourceIdByPartPath, assets: sourceAssets);
            var outputContext = new PptxPartContext(outputSlides[slideIndex], outputIdByPartPath, assets: outputAssets);
            var before = ShapeElements(sourceSlide.Slide!.CommonSlideData!.ShapeTree!);
            var after = ShapeElements(outputRoot.CommonSlideData!.ShapeTree!);
            var afterIds = NativeElementIds(after, requested.Slides[slideIndex].Id);
            var elements = requested.Slides[slideIndex].Elements;
            if (before.Length != elements.Count || after.Length != elements.Count)
                throw new CodecException("presentation_postwrite_topology_changed", $"PPTX slide {slideIndex + 1} element topology changed during source-preserving export.", PartPath(outputSlide));
            for (var elementIndex = 0; elementIndex < elements.Count; elementIndex++)
            {
                var request = elements[elementIndex];
                var binding = request.Source!;
                var changed = !SemanticHash(request).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase);
                if (!changed)
                {
                    if (!HashElement(before[elementIndex]).Equals(HashElement(after[elementIndex]), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "presentation_unchanged_element_modified",
                            $"PPTX slide {slideIndex + 1} unchanged element {elementIndex + 1} was modified during export.",
                            PartPath(outputSlides[slideIndex]));
                    continue;
                }
                if (request.ContentCase == PresentationElement.ContentOneofCase.Opaque)
                {
                    if (!NativeObjectResidualHash(before[elementIndex]).Equals(NativeObjectResidualHash(after[elementIndex]), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "presentation_unmodeled_native_content_changed",
                            $"PPTX slide {slideIndex + 1} edited native object {elementIndex + 1} changed unmodeled native content.",
                            PartPath(outputSlides[slideIndex]));
                    var outputFrame = ReadFrame(after[elementIndex]);
                    if (!ElementName(after[elementIndex], elementIndex).Equals(request.Name, StringComparison.Ordinal) ||
                        outputFrame.Left != request.Opaque.LeftEmu || outputFrame.Top != request.Opaque.TopEmu ||
                        outputFrame.Width != request.Opaque.WidthEmu || outputFrame.Height != request.Opaque.HeightEmu)
                        throw new CodecException(
                            "presentation_postwrite_semantics_mismatch",
                            $"PPTX slide {slideIndex + 1} edited native object {elementIndex + 1} does not match the requested name/frame.",
                            PartPath(outputSlides[slideIndex]));
                    continue;
                }
                if (request.ContentCase == PresentationElement.ContentOneofCase.Group)
                {
                    if (before[elementIndex] is not P.GroupShape beforeGroup || after[elementIndex] is not P.GroupShape afterGroup)
                        throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} edited group {elementIndex + 1} changed native element type.", PartPath(outputSlides[slideIndex]));
                    ValidateGroupOutput(beforeGroup, afterGroup, request, sourceContext, outputContext, afterIds, slideIndex, $"element {elementIndex + 1}");
                    var outputGroupSemantic = ReadElement(afterGroup, slideIndex, elementIndex, outputContext, elementIdsByNativeId: afterIds);
                    if (!SemanticHash(outputGroupSemantic).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException("presentation_postwrite_semantics_mismatch", $"PPTX slide {slideIndex + 1} edited group {elementIndex + 1} does not match requested semantics after export.", PartPath(outputSlides[slideIndex]));
                    continue;
                }
                if (request.ContentCase == PresentationElement.ContentOneofCase.Image)
                {
                    if (before[elementIndex] is not P.Picture beforePicture || after[elementIndex] is not P.Picture afterPicture)
                        throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} edited image {elementIndex + 1} changed native element type.", PartPath(outputSlides[slideIndex]));
                    if (!PictureResidualHash(beforePicture).Equals(PictureResidualHash(afterPicture), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "presentation_unmodeled_picture_content_changed",
                            $"PPTX slide {slideIndex + 1} edited image {elementIndex + 1} changed unmodeled native content.",
                            PartPath(outputSlides[slideIndex]));
                    var outputPictureSemantic = ReadElement(afterPicture, slideIndex, elementIndex, outputContext);
                    if (!SemanticHash(outputPictureSemantic).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "presentation_postwrite_semantics_mismatch",
                            $"PPTX slide {slideIndex + 1} edited image {elementIndex + 1} does not match requested semantics after export.",
                            PartPath(outputSlides[slideIndex]));
                    continue;
                }
                if (request.ContentCase == PresentationElement.ContentOneofCase.Table)
                {
                    if (before[elementIndex] is not P.GraphicFrame beforeTable || after[elementIndex] is not P.GraphicFrame afterTable)
                        throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} edited table {elementIndex + 1} changed native element type.", PartPath(outputSlides[slideIndex]));
                    if (!TableResidualHash(beforeTable).Equals(TableResidualHash(afterTable), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "presentation_unmodeled_table_content_changed",
                            $"PPTX slide {slideIndex + 1} edited table {elementIndex + 1} changed unmodeled native content.",
                            PartPath(outputSlides[slideIndex]));
                    var outputTableSemantic = ReadElement(afterTable, slideIndex, elementIndex, outputContext);
                    if (!SemanticHash(outputTableSemantic).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "presentation_postwrite_semantics_mismatch",
                            $"PPTX slide {slideIndex + 1} edited table {elementIndex + 1} does not match requested semantics after export.",
                            PartPath(outputSlides[slideIndex]));
                    continue;
                }
                if (request.ContentCase == PresentationElement.ContentOneofCase.Connector)
                {
                    if (before[elementIndex] is not P.ConnectionShape beforeConnector || after[elementIndex] is not P.ConnectionShape afterConnector)
                        throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} edited connector {elementIndex + 1} changed native element type.", PartPath(outputSlides[slideIndex]));
                    if (!ConnectorResidualHash(beforeConnector).Equals(ConnectorResidualHash(afterConnector), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException("presentation_unmodeled_connector_content_changed", $"PPTX slide {slideIndex + 1} edited connector {elementIndex + 1} changed unmodeled native content.", PartPath(outputSlides[slideIndex]));
                    var outputConnectorSemantic = ReadElement(afterConnector, slideIndex, elementIndex, outputContext, elementIdsByNativeId: afterIds);
                    if (!SemanticHash(outputConnectorSemantic).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException("presentation_postwrite_semantics_mismatch", $"PPTX slide {slideIndex + 1} edited connector {elementIndex + 1} does not match requested semantics after export.", PartPath(outputSlides[slideIndex]));
                    continue;
                }
                if (request.ContentCase == PresentationElement.ContentOneofCase.Chart)
                {
                    if (before[elementIndex] is not P.GraphicFrame beforeChart || after[elementIndex] is not P.GraphicFrame afterChart)
                        throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} edited chart {elementIndex + 1} changed native element type.", PartPath(outputSlides[slideIndex]));
                    if (!ChartFrameResidualHash(beforeChart).Equals(ChartFrameResidualHash(afterChart), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException("presentation_unmodeled_chart_frame_changed", $"PPTX slide {slideIndex + 1} edited chart {elementIndex + 1} changed unmodeled frame content.", PartPath(outputSlides[slideIndex]));
                    var outputChartSemantic = ReadElement(afterChart, slideIndex, elementIndex, outputContext, elementIdsByNativeId: afterIds);
                    if (!SemanticHash(outputChartSemantic).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException("presentation_postwrite_semantics_mismatch", $"PPTX slide {slideIndex + 1} edited chart {elementIndex + 1} does not match requested semantics after export.", PartPath(outputSlides[slideIndex]));
                    continue;
                }
                if (before[elementIndex] is not P.Shape beforeShape || after[elementIndex] is not P.Shape afterShape)
                    throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} edited element {elementIndex + 1} changed native element type.", PartPath(outputSlides[slideIndex]));
                if (!ShapeResidualHash(beforeShape, sourceContext).Equals(ShapeResidualHash(afterShape, outputContext), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_unmodeled_shape_content_changed",
                        $"PPTX slide {slideIndex + 1} edited shape {elementIndex + 1} changed unmodeled native content.",
                        PartPath(outputSlides[slideIndex]));
                var outputSemantic = ReadElement(afterShape, slideIndex, elementIndex, outputContext, elementIdsByNativeId: afterIds);
                if (!SemanticHash(outputSemantic).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_postwrite_semantics_mismatch",
                        $"PPTX slide {slideIndex + 1} edited shape {elementIndex + 1} does not match requested semantics after export.",
                        PartPath(outputSlides[slideIndex]));
            }
        }
    }

    private static void ValidateClonedRunHyperlinkGraph(SlidePart source, SlidePart clone, int targetIndex)
    {
        var sourceExternal = source.HyperlinkRelationships
            .ToDictionary(
                relationship => relationship.Id,
                relationship => (relationship.Uri.OriginalString, relationship.IsExternal),
                StringComparer.Ordinal);
        var cloneExternal = clone.HyperlinkRelationships
            .ToDictionary(
                relationship => relationship.Id,
                relationship => (relationship.Uri.OriginalString, relationship.IsExternal),
                StringComparer.Ordinal);
        var sourceSlides = source.Parts
            .Where(pair => pair.OpenXmlPart is SlidePart)
            .ToDictionary(
                pair => pair.RelationshipId,
                pair => PartPath(pair.OpenXmlPart),
                StringComparer.Ordinal);
        var cloneSlides = clone.Parts
            .Where(pair => pair.OpenXmlPart is SlidePart)
            .ToDictionary(
                pair => pair.RelationshipId,
                pair => PartPath(pair.OpenXmlPart),
                StringComparer.Ordinal);

        if (sourceExternal.Count != cloneExternal.Count ||
            sourceExternal.Any(pair => !cloneExternal.TryGetValue(pair.Key, out var actual) || actual != pair.Value) ||
            sourceSlides.Count != cloneSlides.Count ||
            sourceSlides.Any(pair => !cloneSlides.TryGetValue(pair.Key, out var actual) || !actual.Equals(pair.Value, StringComparison.OrdinalIgnoreCase)) ||
            clone.ExternalRelationships.Any() || clone.DataPartReferenceRelationships.Any())
            throw new CodecException(
                "presentation_postwrite_clone_hyperlink_mismatch",
                $"PPTX clone {targetIndex + 1} does not retain the exact bounded run-hyperlink relationship graph.",
                PartPath(clone));
    }

    private static void ValidateClonedSpeakerNotes(SlidePart source, SlidePart clone, int targetIndex)
    {
        var sourceNotes = source.NotesSlidePart;
        var cloneNotes = clone.NotesSlidePart;
        if (sourceNotes is null)
        {
            if (cloneNotes is not null)
                throw new CodecException("presentation_postwrite_clone_notes_mismatch", $"PPTX clone {targetIndex + 1} unexpectedly added a notes part.", PartPath(clone));
            return;
        }
        if (cloneNotes?.NotesSlide is null || sourceNotes.NotesSlide is null ||
            PartPath(sourceNotes).Equals(PartPath(cloneNotes), StringComparison.OrdinalIgnoreCase) ||
            !PartBytes(sourceNotes).SequenceEqual(PartBytes(cloneNotes)))
            throw new CodecException("presentation_postwrite_clone_notes_mismatch", $"PPTX clone {targetIndex + 1} is not an independent exact source notes copy.", PartPath(clone));
        var sourceMaster = sourceNotes.NotesMasterPart;
        var cloneMaster = cloneNotes.NotesMasterPart;
        if (sourceMaster is null || cloneMaster is null ||
            !PartPath(sourceMaster).Equals(PartPath(cloneMaster), StringComparison.OrdinalIgnoreCase) ||
            source.GetIdOfPart(sourceNotes) != clone.GetIdOfPart(cloneNotes) ||
            sourceNotes.GetIdOfPart(sourceMaster) != cloneNotes.GetIdOfPart(cloneMaster) ||
            sourceNotes.GetIdOfPart(source) != cloneNotes.GetIdOfPart(clone) ||
            cloneNotes.Parts.Count() != 2 || cloneNotes.ExternalRelationships.Any() || cloneNotes.HyperlinkRelationships.Any() || cloneNotes.DataPartReferenceRelationships.Any())
            throw new CodecException("presentation_postwrite_clone_notes_mismatch", $"PPTX clone {targetIndex + 1} does not retain the exact bounded notes relationship graph.", PartPath(cloneNotes));
    }

    private static void ValidateClonedLegacyComments(
        PresentationPart sourcePresentation,
        PresentationPart outputPresentation,
        SlidePart source,
        SlidePart clone,
        int targetIndex)
    {
        var sourceComments = source.SlideCommentsPart;
        var cloneComments = clone.SlideCommentsPart;
        if (sourceComments is null)
        {
            if (cloneComments is not null)
                throw new CodecException("presentation_postwrite_clone_comments_mismatch", $"PPTX clone {targetIndex + 1} unexpectedly added a comments part.", PartPath(clone));
            return;
        }
        if (cloneComments?.CommentList is null || sourceComments.CommentList is null ||
            PartPath(sourceComments).Equals(PartPath(cloneComments), StringComparison.OrdinalIgnoreCase) ||
            !PartBytes(sourceComments).SequenceEqual(PartBytes(cloneComments)) ||
            source.GetIdOfPart(sourceComments) != clone.GetIdOfPart(cloneComments) ||
            cloneComments.Parts.Any() || cloneComments.ExternalRelationships.Any() ||
            cloneComments.HyperlinkRelationships.Any() || cloneComments.DataPartReferenceRelationships.Any())
            throw new CodecException("presentation_postwrite_clone_comments_mismatch", $"PPTX clone {targetIndex + 1} is not an independent exact source legacy-comments copy.", PartPath(clone));

        // p:cm/@authorId resolves through this shared catalog rather than a
        // comment-part relationship. Prove that the clone still references
        // the one untouched catalog in the output package.
        var sourceAuthors = sourcePresentation.CommentAuthorsPart;
        var outputAuthors = outputPresentation.CommentAuthorsPart;
        if (sourceAuthors is null || outputAuthors is null ||
            !PartPath(sourceAuthors).Equals(PartPath(outputAuthors), StringComparison.OrdinalIgnoreCase) ||
            !PartBytes(sourceAuthors).SequenceEqual(PartBytes(outputAuthors)) ||
            sourcePresentation.GetIdOfPart(sourceAuthors) != outputPresentation.GetIdOfPart(outputAuthors) ||
            outputAuthors.Parts.Any() || outputAuthors.ExternalRelationships.Any() ||
            outputAuthors.HyperlinkRelationships.Any() || outputAuthors.DataPartReferenceRelationships.Any())
            throw new CodecException("presentation_postwrite_clone_comments_mismatch", $"PPTX clone {targetIndex + 1} does not retain the bounded shared legacy-comment author catalog.", PartPath(cloneComments));
    }

    private static void ValidateGroupOutput(
        P.GroupShape before,
        P.GroupShape after,
        PresentationElement request,
        PptxPartContext sourceContext,
        PptxPartContext outputContext,
        IReadOnlyDictionary<uint, string> afterIds,
        int slideIndex,
        string location)
    {
        if (!GroupShellResidualHash(before).Equals(GroupShellResidualHash(after), StringComparison.OrdinalIgnoreCase))
            throw new CodecException("presentation_unmodeled_group_content_changed", $"PPTX slide {slideIndex + 1} {location} changed unmodeled group-shell content.", PartPath(outputContext.Owner));
        var beforeChildren = GroupElements(before);
        var afterChildren = GroupElements(after);
        if (beforeChildren.Length != request.Group.Children.Count || afterChildren.Length != request.Group.Children.Count)
            throw new CodecException("presentation_postwrite_topology_changed", $"PPTX slide {slideIndex + 1} {location} group topology changed during export.", PartPath(outputContext.Owner));

        for (var index = 0; index < request.Group.Children.Count; index++)
        {
            var child = request.Group.Children[index];
            var binding = child.Source ?? throw new CodecException("missing_presentation_element_binding", $"PPTX slide {slideIndex + 1} {location} child {index + 1} is missing its source binding.", PartPath(outputContext.Owner));
            var changed = !SemanticHash(child).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase);
            if (!changed)
            {
                if (!HashElement(beforeChildren[index]).Equals(HashElement(afterChildren[index]), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_unchanged_element_modified", $"PPTX slide {slideIndex + 1} {location} unchanged child {index + 1} was modified during export.", PartPath(outputContext.Owner));
                continue;
            }

            if (child.ContentCase == PresentationElement.ContentOneofCase.Group)
            {
                if (beforeChildren[index] is not P.GroupShape beforeGroup || afterChildren[index] is not P.GroupShape afterGroup)
                    throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} {location} child group {index + 1} changed native element type.", PartPath(outputContext.Owner));
                ValidateGroupOutput(beforeGroup, afterGroup, child, sourceContext, outputContext, afterIds, slideIndex, $"{location} child {index + 1}");
            }
            else if (child.ContentCase == PresentationElement.ContentOneofCase.Image)
            {
                if (beforeChildren[index] is not P.Picture beforePicture || afterChildren[index] is not P.Picture afterPicture ||
                    !PictureResidualHash(beforePicture).Equals(PictureResidualHash(afterPicture), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_unmodeled_picture_content_changed", $"PPTX slide {slideIndex + 1} {location} child image {index + 1} changed unmodeled content.", PartPath(outputContext.Owner));
            }
            else if (child.ContentCase == PresentationElement.ContentOneofCase.Table)
            {
                if (beforeChildren[index] is not P.GraphicFrame beforeTable || afterChildren[index] is not P.GraphicFrame afterTable ||
                    !TableResidualHash(beforeTable).Equals(TableResidualHash(afterTable), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_unmodeled_table_content_changed", $"PPTX slide {slideIndex + 1} {location} child table {index + 1} changed unmodeled content.", PartPath(outputContext.Owner));
            }
            else if (child.ContentCase == PresentationElement.ContentOneofCase.Connector)
            {
                if (beforeChildren[index] is not P.ConnectionShape beforeConnector || afterChildren[index] is not P.ConnectionShape afterConnector ||
                    !ConnectorResidualHash(beforeConnector).Equals(ConnectorResidualHash(afterConnector), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_unmodeled_connector_content_changed", $"PPTX slide {slideIndex + 1} {location} child connector {index + 1} changed unmodeled content.", PartPath(outputContext.Owner));
            }
            else if (child.ContentCase == PresentationElement.ContentOneofCase.Chart)
            {
                if (beforeChildren[index] is not P.GraphicFrame beforeChart || afterChildren[index] is not P.GraphicFrame afterChart ||
                    !ChartFrameResidualHash(beforeChart).Equals(ChartFrameResidualHash(afterChart), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_unmodeled_chart_frame_changed", $"PPTX slide {slideIndex + 1} {location} child chart {index + 1} changed unmodeled frame content.", PartPath(outputContext.Owner));
            }
            else
            {
                if (beforeChildren[index] is not P.Shape beforeShape || afterChildren[index] is not P.Shape afterShape ||
                    !ShapeResidualHash(beforeShape, sourceContext).Equals(ShapeResidualHash(afterShape, outputContext), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_unmodeled_shape_content_changed", $"PPTX slide {slideIndex + 1} {location} child shape {index + 1} changed unmodeled content.", PartPath(outputContext.Owner));
            }

            var outputSemantic = ReadElement(afterChildren[index], request.Id, index, outputContext, elementIdsByNativeId: afterIds);
            if (!SemanticHash(outputSemantic).Equals(SemanticHash(child), StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_postwrite_semantics_mismatch", $"PPTX slide {slideIndex + 1} {location} child {index + 1} does not match requested semantics after export.", PartPath(outputContext.Owner));
        }
    }

    private static string GroupShellResidualHash(P.GroupShape source)
    {
        var clone = (P.GroupShape)source.CloneNode(true);
        foreach (var child in GroupElements(clone)) child.Remove();
        if (clone.NonVisualGroupShapeProperties?.NonVisualDrawingProperties is { } nonVisual) nonVisual.Name = string.Empty;
        if (clone.GroupShapeProperties?.GetFirstChild<A.TransformGroup>() is { } transform)
        {
            transform.Offset!.X = 0L;
            transform.Offset.Y = 0L;
            transform.Extents!.Cx = 1L;
            transform.Extents.Cy = 1L;
            transform.ChildOffset!.X = 0L;
            transform.ChildOffset.Y = 0L;
            transform.ChildExtents!.Cx = 1L;
            transform.ChildExtents.Cy = 1L;
        }
        return HashElement(clone);
    }

    private static void ValidatePreservedMasterAndLayoutContent(byte[] sourceBytes, byte[] outputBytes, PresentationArtifact requested, EffectiveCodecLimits limits)
    {
        using var sourceStream = new MemoryStream(sourceBytes, writable: false);
        using var outputStream = new MemoryStream(outputBytes, writable: false);
        using var sourcePackage = PresentationDocument.Open(sourceStream, isEditable: false);
        using var outputPackage = PresentationDocument.Open(outputStream, isEditable: false);
        var sourcePresentationPart = sourcePackage.PresentationPart ??
            throw new CodecException("missing_presentation_part", "PPTX source package has no Presentation part.", "ppt/presentation.xml");
        var outputPresentationPart = outputPackage.PresentationPart ??
            throw new CodecException("missing_presentation_part", "PPTX output package has no Presentation part.", "ppt/presentation.xml");
        var sourceGraph = ReadMasterGraph(sourcePresentationPart);
        var outputGraph = ReadMasterGraph(outputPresentationPart);
        if (sourceGraph.Length != requested.Masters.Count || outputGraph.Length != requested.Masters.Count)
            throw new CodecException("presentation_postwrite_master_topology_changed", "PPTX master topology changed during source-preserving export.");
        var sourceTargets = BindSourcePreservingSlides(
            sourcePresentationPart,
            sourcePresentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>().ToArray() ?? [],
            requested.Slides);
        var outputSlides = OrderedSlideParts(outputPackage);
        if (outputSlides.Length != sourceTargets.Length)
            throw new CodecException("presentation_postwrite_topology_changed", "PPTX slide order does not match the requested source-bound order.", "ppt/presentation.xml");
        for (var targetIndex = 0; targetIndex < sourceTargets.Length; targetIndex++)
        {
            var target = sourceTargets[targetIndex];
            if (!target.IsClone && !PartPath(outputSlides[targetIndex]).Equals(PartPath(target.Source.Part), StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_postwrite_topology_changed", "PPTX slide order does not match the requested source-bound order.", "ppt/presentation.xml");
        }
        var retainedTargets = sourceTargets.Where(target => !target.IsClone).ToArray();
        var sourceSlideMap = retainedTargets.Select(target => (Path: PartPath(target.Source.Part), Id: target.Target.Id))
            .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
        var outputSlideMap = retainedTargets.Select(target => (Path: PartPath(target.Source.Part), Id: target.Target.Id))
            .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
        var sourceAssets = new PptxAssetCatalog([], limits);
        var outputAssets = new PptxAssetCatalog([], limits);
        for (var masterIndex = 0; masterIndex < requested.Masters.Count; masterIndex++)
        {
            var before = sourceGraph[masterIndex].Part.SlideMaster ??
                throw new CodecException("missing_slide_master_root", $"PPTX source master {masterIndex + 1} has no root.");
            var after = outputGraph[masterIndex].Part.SlideMaster ??
                throw new CodecException("missing_slide_master_root", $"PPTX output master {masterIndex + 1} has no root.");
            var sourceContext = new PptxPartContext(sourceGraph[masterIndex].Part, sourceSlideMap, assets: sourceAssets);
            var outputContext = new PptxPartContext(outputGraph[masterIndex].Part, outputSlideMap, assets: outputAssets);
            if (!MasterResidualHash(before, sourceContext).Equals(MasterResidualHash(after, outputContext), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_unmodeled_master_content_changed",
                    $"PPTX master {masterIndex + 1} edit changed unmodeled native content.",
                    PartPath(outputGraph[masterIndex].Part));
            var outputStyles = PptxMasterTextStylesCodec.Read(after, outputContext);
            if (!MasterTextStylesSemanticHash(outputStyles).Equals(MasterTextStylesSemanticHash(requested.Masters[masterIndex].TextStyles), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_postwrite_master_semantics_mismatch",
                    $"PPTX master {masterIndex + 1} text styles do not match requested semantics after export.",
                    PartPath(outputGraph[masterIndex].Part));
            if (!BackgroundSemanticHash(PptxBackgroundCodec.Read(after.CommonSlideData)).Equals(BackgroundSemanticHash(requested.Masters[masterIndex].Background), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_postwrite_master_background_mismatch",
                    $"PPTX master {masterIndex + 1} background does not match requested semantics after export.",
                    PartPath(outputGraph[masterIndex].Part));
            ValidatePlaceholderOutput(
                before.CommonSlideData?.ShapeTree,
                after.CommonSlideData?.ShapeTree,
                requested.Masters[masterIndex].Placeholders,
                requested.Masters[masterIndex].Id,
                sourceContext,
                outputContext,
                PartPath(outputGraph[masterIndex].Part));
        }
        var sourceLayouts = sourceGraph.SelectMany(master => master.Layouts).ToArray();
        var outputLayouts = outputGraph.SelectMany(master => master.Layouts).ToArray();
        if (sourceLayouts.Length != requested.Layouts.Count || outputLayouts.Length != requested.Layouts.Count)
            throw new CodecException("presentation_postwrite_layout_topology_changed", "PPTX layout topology changed during source-preserving export.");
        for (var layoutIndex = 0; layoutIndex < requested.Layouts.Count; layoutIndex++)
        {
            var before = sourceLayouts[layoutIndex].Part.SlideLayout ??
                throw new CodecException("missing_slide_layout_root", $"PPTX source layout {layoutIndex + 1} has no root.");
            var after = outputLayouts[layoutIndex].Part.SlideLayout ??
                throw new CodecException("missing_slide_layout_root", $"PPTX output layout {layoutIndex + 1} has no root.");
            var sourceContext = new PptxPartContext(sourceLayouts[layoutIndex].Part, sourceSlideMap, assets: sourceAssets);
            var outputContext = new PptxPartContext(outputLayouts[layoutIndex].Part, outputSlideMap, assets: outputAssets);
            if (!LayoutResidualHash(before, sourceContext).Equals(LayoutResidualHash(after, outputContext), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_unmodeled_layout_content_changed",
                    $"PPTX layout {layoutIndex + 1} edit changed unmodeled native content.",
                    PartPath(outputLayouts[layoutIndex].Part));
            if (!BackgroundSemanticHash(PptxBackgroundCodec.Read(after.CommonSlideData)).Equals(BackgroundSemanticHash(requested.Layouts[layoutIndex].Background), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_postwrite_layout_background_mismatch",
                    $"PPTX layout {layoutIndex + 1} background does not match requested semantics after export.",
                    PartPath(outputLayouts[layoutIndex].Part));
            ValidatePlaceholderOutput(
                before.CommonSlideData?.ShapeTree,
                after.CommonSlideData?.ShapeTree,
                requested.Layouts[layoutIndex].Placeholders,
                requested.Layouts[layoutIndex].Id,
                sourceContext,
                outputContext,
                PartPath(outputLayouts[layoutIndex].Part));
        }
    }

    private static void ValidatePlaceholderOutput(
        P.ShapeTree? sourceTree,
        P.ShapeTree? outputTree,
        IList<PresentationPlaceholder> requested,
        string ownerId,
        PptxPartContext sourceContext,
        PptxPartContext outputContext,
        string partPath)
    {
        if (sourceTree is null || outputTree is null)
            throw new CodecException("missing_shape_tree", $"Presentation owner {ownerId} has no shape tree.", partPath);
        var before = PptxPlaceholderCodec.Read(sourceTree, ownerId, sourceContext);
        var after = PptxPlaceholderCodec.Read(outputTree, ownerId, outputContext);
        if (before.Count != requested.Count || after.Count != requested.Count)
            throw new CodecException("presentation_postwrite_placeholder_topology_changed", $"Presentation owner {ownerId} placeholder topology changed during export.", partPath);
        for (var index = 0; index < requested.Count; index++)
        {
            var request = requested[index];
            var binding = request.Source ?? throw new CodecException("missing_presentation_placeholder_binding", $"Presentation placeholder {index + 1} under {ownerId} is missing its source binding.", partPath);
            var changed = !PptxPlaceholderCodec.SemanticHash(request).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase);
            if (!changed)
            {
                var sourceShape = PptxPlaceholderCodec.BoundShape(sourceTree, before[index]);
                var outputShape = PptxPlaceholderCodec.BoundShape(outputTree, after[index]);
                if (sourceShape is null || outputShape is null ||
                    !PptxPlaceholderCodec.ElementHash(sourceShape).Equals(PptxPlaceholderCodec.ElementHash(outputShape), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_unchanged_placeholder_modified", $"Presentation placeholder {index + 1} under {ownerId} was modified during export.", partPath);
            }
            if (after[index].Id != request.Id ||
                !PptxPlaceholderCodec.SemanticHash(after[index]).Equals(PptxPlaceholderCodec.SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_postwrite_placeholder_semantics_mismatch", $"Presentation placeholder {index + 1} under {ownerId} does not match requested semantics after export.", partPath);
        }
    }

    private static SlidePart[] OrderedSlideParts(PresentationDocument package)
    {
        var presentationPart = package.PresentationPart ?? throw new CodecException("missing_presentation_part", "PPTX package has no Presentation part.", "ppt/presentation.xml");
        return ResolveSlideParts(presentationPart, presentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>() ?? []);
    }

    private static PptxMasterGraphEntry[] ReadMasterGraph(PresentationPart presentationPart)
    {
        var masterIds = presentationPart.Presentation?.SlideMasterIdList?.Elements<P.SlideMasterId>().ToArray() ?? [];
        return masterIds.Select((masterId, masterIndex) =>
        {
            var relationshipId = masterId.RelationshipId?.Value ?? string.Empty;
            var masterPart = presentationPart.GetPartById(relationshipId) as SlideMasterPart ??
                throw new CodecException("missing_slide_master_part", $"Presentation master {masterIndex + 1} has an unresolved relationship.", "ppt/presentation.xml");
            var layoutIds = masterPart.SlideMaster?.SlideLayoutIdList?.Elements<P.SlideLayoutId>().ToArray() ?? [];
            var layouts = layoutIds.Select((layoutId, layoutIndex) =>
            {
                var layoutRelationshipId = layoutId.RelationshipId?.Value ?? string.Empty;
                var layoutPart = masterPart.GetPartById(layoutRelationshipId) as SlideLayoutPart ??
                    throw new CodecException("missing_slide_layout_part", $"Presentation layout {layoutIndex + 1} under master {masterIndex + 1} has an unresolved relationship.", PartPath(masterPart));
                return new PptxLayoutGraphEntry(
                    layoutIndex,
                    $"presentation/master/{masterIndex + 1}/layout/{layoutIndex + 1}",
                    layoutRelationshipId,
                    layoutPart);
            }).ToArray();
            return new PptxMasterGraphEntry(
                masterIndex,
                $"presentation/master/{masterIndex + 1}",
                relationshipId,
                masterPart,
                layouts);
        }).ToArray();
    }

    private static string LayoutTypeName(P.SlideLayout source)
    {
        // `type` is a typed PresentationML attribute. Reading it through the
        // generic attribute accessor with an empty namespace is rejected by
        // Open XML SDK 3.x for typed elements, even though it is a valid
        // unqualified attribute in the serialized package. Use the generated
        // property so imported layouts from Office-authored templates remain
        // readable.
        var value = source.Type?.InnerText;
        return string.IsNullOrWhiteSpace(value) ? "custom" : value;
    }

    private static SlidePart[] ResolveSlideParts(PresentationPart presentationPart, IEnumerable<P.SlideId> slideIds) =>
        slideIds.Select(slideId => presentationPart.GetPartById(slideId.RelationshipId?.Value ?? string.Empty) as SlidePart ??
            throw new CodecException("missing_slide_part", "PPTX presentation contains an unresolved slide relationship.", "ppt/presentation.xml"))
        .ToArray();

    // Imported source topology keeps ordinary bindings and clone requests
    // separate. A clone is a new SlidePart with a verified origin, never a
    // second presentation reference to the same source part.
    private static PptxTargetSlideEntry[] BindSourcePreservingSlides(
        PresentationPart presentationPart,
        IReadOnlyList<P.SlideId> sourceSlideIds,
        IReadOnlyList<PresentationSlide> requested)
    {
        if (requested.Count == 0)
            throw new CodecException(
                "presentation_topology_changed",
                "Source-preserving PPTX export must retain at least one source slide.",
                "ppt/presentation.xml");
        var sourceParts = ResolveSlideParts(presentationPart, sourceSlideIds);
        var sourceSlides = sourceSlideIds
            .Select((slideId, index) => new PptxSourceSlideEntry(
                index,
                slideId,
                slideId.RelationshipId?.Value ?? string.Empty,
                sourceParts[index]))
            .ToArray();
        var targets = new PptxTargetSlideEntry[requested.Count];
        var seenSourceParts = new HashSet<SlidePart>();
        var seenCloneSourceParts = new HashSet<SlidePart>();
        for (var targetIndex = 0; targetIndex < requested.Count; targetIndex++)
        {
            var target = requested[targetIndex];
            var isClone = target.CloneSource is not null;
            if ((target.Source is null) == !isClone)
                throw new CodecException(
                    "presentation_slide_binding_mismatch",
                    $"Presentation slide {targetIndex + 1} must carry exactly one of source or clone_source.",
                    "ppt/presentation.xml");
            var binding = target.Source ?? target.CloneSource!;
            if (binding.SlideIndex >= sourceSlides.Length)
                throw new CodecException(
                    "presentation_slide_binding_mismatch",
                    $"Presentation slide {targetIndex + 1} references source slide {binding.SlideIndex + 1}, which does not exist.",
                    "ppt/presentation.xml");
            var source = sourceSlides[binding.SlideIndex];
            var sourceRoot = source.Part.Slide ??
                throw new CodecException("missing_slide_root", $"Presentation source slide {source.Index + 1} has no slide root.", PartPath(source.Part));
            if (!binding.PartPath.Equals(PartPath(source.Part), StringComparison.OrdinalIgnoreCase) ||
                !binding.RelationshipId.Equals(source.RelationshipId, StringComparison.Ordinal) ||
                !binding.SlideXmlSha256.Equals(HashElement(sourceRoot), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_slide_binding_mismatch",
                    $"Presentation slide {targetIndex + 1} does not match its hash-bound source slide.",
                    PartPath(source.Part));
            if (!isClone && !seenSourceParts.Add(source.Part))
                throw new CodecException(
                    "presentation_topology_changed",
                    "Source-preserving PPTX export cannot bind more than one ordinary target to a source SlidePart.",
                    "ppt/presentation.xml");
            if (isClone && !seenCloneSourceParts.Add(source.Part))
                throw new CodecException(
                    "unsupported_presentation_slide_clone",
                    "The bounded source-preserving PPTX clone profile permits only one pending clone per source SlidePart.",
                    PartPath(source.Part));
            targets[targetIndex] = new PptxTargetSlideEntry(targetIndex, target, source, isClone);
        }
        return targets;
    }

    private static void DeleteUnrequestedSourceSlides(
        PresentationPart presentationPart,
        IReadOnlyList<P.SlideId> sourceSlideIds,
        IReadOnlyList<PptxTargetSlideEntry> targets,
        OpaqueOpcGraph opaque,
        ISet<string> changedParts,
        ISet<string> removedSourceSlidePartPaths)
    {
        var sourceParts = ResolveSlideParts(presentationPart, sourceSlideIds);
        var retainedParts = targets.Where(target => !target.IsClone).Select(target => target.Source.Part).ToHashSet();
        var removed = sourceSlideIds
            .Select((slideId, index) => new PptxSourceSlideEntry(
                index,
                slideId,
                slideId.RelationshipId?.Value ?? string.Empty,
                sourceParts[index]))
            .Where(source => !retainedParts.Contains(source.Part))
            .ToArray();
        if (removed.Length == 0) return;

        foreach (var source in removed)
            AssertSourceSlideCanBeDeleted(presentationPart, source, opaque);

        foreach (var source in removed)
        {
            var slidePath = PartPath(source.Part);
            var relationshipPath = RelationshipPartPath(source.Part);
            presentationPart.DeletePart(source.Part);
            changedParts.Add(slidePath);
            changedParts.Add(relationshipPath);
            removedSourceSlidePartPaths.Add(slidePath);
        }
        changedParts.Add(PartPath(presentationPart));
        changedParts.Add(RelationshipPartPath(presentationPart));
        changedParts.Add("[Content_Types].xml");
    }

    private static void AssertSourceSlideCanBeDeleted(
        PresentationPart presentationPart,
        PptxSourceSlideEntry source,
        OpaqueOpcGraph opaque)
    {
        var slidePath = PartPath(source.Part);
        var childParts = source.Part.Parts.ToArray();
        var layoutRelationshipCount = childParts.Count(pair => pair.OpenXmlPart is SlideLayoutPart);
        var unsafeChildren = childParts
            .Where(pair => pair.OpenXmlPart is not SlideLayoutPart)
            .Select(pair => pair.OpenXmlPart.RelationshipType)
            .Distinct(StringComparer.Ordinal)
            .Take(4)
            .ToArray();
        if (layoutRelationshipCount != 1 ||
            unsafeChildren.Length > 0 ||
            source.Part.ExternalRelationships.Any() ||
            source.Part.HyperlinkRelationships.Any() ||
            source.Part.DataPartReferenceRelationships.Any())
            throw UnsupportedSourceSlideDelete(
                source,
                "it owns media, notes, comments, hyperlinks, data parts, or another non-layout relationship");

        var inbound = opaque.PackageRelationships
            .Where(relationship => !relationship.TargetMode.Equals("External", StringComparison.OrdinalIgnoreCase))
            .Where(relationship => PptxNativeObjectCatalog.ResolveTarget(relationship.SourcePath, relationship.Target)
                .Equals(slidePath, StringComparison.OrdinalIgnoreCase))
            .Take(4)
            .ToArray();
        if (inbound.Length > 0)
            throw UnsupportedSourceSlideDelete(
                source,
                $"it is referenced by {string.Join(", ", inbound.Select(relationship => relationship.SourcePath))}");

        var root = presentationPart.Presentation ??
            throw new CodecException("missing_presentation_root", "PPTX package has no Presentation root.", "ppt/presentation.xml");
        if (root.ChildElements.Any(element => element.LocalName is "custShowLst" or "sectionLst" or "extLst"))
            throw UnsupportedSourceSlideDelete(
                source,
                "the presentation contains custom shows, sections, or extension data that can retain slide identity");

        var relationshipReferencesOutsideSlideIds = root
            .Descendants()
            .Where(element => element is not P.SlideId)
            .SelectMany(element => element.GetAttributes())
            .Any(attribute => (attribute.NamespaceUri is "http://schemas.openxmlformats.org/officeDocument/2006/relationships" or "http://purl.oclc.org/ooxml/officeDocument/relationships") &&
                              string.Equals(attribute.Value, source.RelationshipId, StringComparison.Ordinal));
        if (relationshipReferencesOutsideSlideIds)
            throw UnsupportedSourceSlideDelete(
                source,
                "a presentation-level relationship outside p:sldIdLst still references it");
    }

    private static CodecException UnsupportedSourceSlideDelete(PptxSourceSlideEntry source, string reason) =>
        new(
            "unsupported_presentation_slide_delete",
            $"Source-preserving PPTX deletion is limited to an isolated layout-only slide; slide {source.Index + 1} cannot be deleted because {reason}. Use an explicit OPC graph delete operation for this package.",
            PartPath(source.Part));

    private static void CloneRequestedSourceSlides(
        PresentationPart presentationPart,
        IReadOnlyList<PptxTargetSlideEntry> targets,
        IReadOnlyDictionary<string, string> layoutIdByPartPath,
        IReadOnlyDictionary<string, string> slideIdByPartPath,
        PptxAssetCatalog assetCatalog,
        ISet<string> changedParts,
        ISet<string> addedRelationshipIds,
        ISet<string> addedPartPaths)
    {
        var cloneTargets = targets.Where(target => target.IsClone).ToArray();
        if (cloneTargets.Length == 0) return;
        var retainedSlideParts = targets
            .Where(target => !target.IsClone)
            .Select(target => target.Source.Part)
            .ToHashSet();
        var root = presentationPart.Presentation ??
            throw new CodecException("missing_presentation_root", "PPTX package has no Presentation root.", "ppt/presentation.xml");
        var slideIdList = root.SlideIdList ??
            throw new CodecException("missing_slide_id_list", "PPTX package has no slide ID list.", "ppt/presentation.xml");
        var nextSlideId = slideIdList.Elements<P.SlideId>()
            .Select(slideId => slideId.Id?.Value ?? 255U)
            .DefaultIfEmpty(255U)
            .Max();

        foreach (var target in cloneTargets)
        {
            var leaf = AssertSourceSlideCanBeCloned(presentationPart, target, layoutIdByPartPath, slideIdByPartPath, retainedSlideParts, assetCatalog);
            var sourcePart = target.Source.Part;
            var sourceRoot = sourcePart.Slide ??
                throw new CodecException("missing_slide_root", $"Presentation source slide {target.Source.Index + 1} has no slide root.", PartPath(sourcePart));
            var clonePart = presentationPart.AddNewPart<SlidePart>();
            clonePart.AddPart(leaf.LayoutPart, leaf.LayoutRelationshipId);
            // The source XML keeps its r:embed values verbatim. Copy each
            // verified internal relationship onto the new SlidePart so those
            // bindings resolve to the same immutable media Parts without
            // duplicating or mutating their bytes.
            foreach (var image in leaf.Images)
            {
                clonePart.AddPart(image.Part, image.RelationshipId);
                addedRelationshipIds.Add($"{PartPath(clonePart)}\0{image.RelationshipId}");
            }
            foreach (var hyperlink in leaf.ExternalHyperlinks)
            {
                clonePart.AddHyperlinkRelationship(hyperlink.Uri, true, hyperlink.RelationshipId);
                addedRelationshipIds.Add($"{PartPath(clonePart)}\0{hyperlink.RelationshipId}");
            }
            foreach (var hyperlink in leaf.SlideHyperlinks)
            {
                clonePart.AddPart(hyperlink.Part, hyperlink.RelationshipId);
                addedRelationshipIds.Add($"{PartPath(clonePart)}\0{hyperlink.RelationshipId}");
            }
            if (leaf.Notes is { } notes)
            {
                var cloneNotes = clonePart.AddNewPart<NotesSlidePart>(notes.SlideRelationshipId);
                cloneNotes.AddPart(notes.MasterPart, notes.MasterRelationshipId);
                cloneNotes.AddPart(clonePart, notes.SlideBackReferenceRelationshipId);
                // The NotesSlide is an opaque-preserved leaf here. Assigning
                // its DOM would normalize namespace declarations on save, so
                // copy the OPC payload bytes after wiring its two verified
                // relationships. It remains untouched until export/reimport.
                CopyPartBytes(notes.Part, cloneNotes);
                addedRelationshipIds.Add($"{PartPath(clonePart)}\0{notes.SlideRelationshipId}");
                addedRelationshipIds.Add($"{PartPath(cloneNotes)}\0{notes.MasterRelationshipId}");
                addedRelationshipIds.Add($"{PartPath(cloneNotes)}\0{notes.SlideBackReferenceRelationshipId}");
                changedParts.Add(PartPath(cloneNotes));
                changedParts.Add(RelationshipPartPath(cloneNotes));
                addedPartPaths.Add(PartPath(cloneNotes));
                addedPartPaths.Add(RelationshipPartPath(cloneNotes));
            }
            if (leaf.LegacyComments is { } comments)
            {
                // Legacy comment XML is another opaque-preserved leaf. Its
                // author IDs resolve through the already-verified immutable
                // presentation catalog, so copy the bytes but do not create
                // a second catalog or rewrite either source part.
                if (!ReferenceEquals(presentationPart.CommentAuthorsPart, comments.AuthorsPart) ||
                    presentationPart.GetIdOfPart(comments.AuthorsPart) != comments.AuthorsRelationshipId)
                    throw new CodecException("presentation_slide_clone_binding_mismatch", "The shared legacy-comment author catalog changed after clone preflight.", PartPath(comments.AuthorsPart));
                var cloneComments = clonePart.AddNewPart<SlideCommentsPart>(comments.SlideRelationshipId);
                CopyPartBytes(comments.Part, cloneComments);
                addedRelationshipIds.Add($"{PartPath(clonePart)}\0{comments.SlideRelationshipId}");
                changedParts.Add(PartPath(cloneComments));
                addedPartPaths.Add(PartPath(cloneComments));
            }
            clonePart.Slide = (P.Slide)sourceRoot.CloneNode(true);
            clonePart.Slide.Save();
            if (nextSlideId == uint.MaxValue)
                throw new CodecException("presentation_slide_id_exhausted", "PPTX cannot allocate another 32-bit slide identifier.", "ppt/presentation.xml");
            nextSlideId++;
            target.OutputPart = clonePart;
            target.OutputSlideId = new P.SlideId
            {
                Id = nextSlideId,
                RelationshipId = presentationPart.GetIdOfPart(clonePart),
            };
            // Clone-to-layout, clone-to-image, clone-to-run-hyperlink,
            // clone-to-notes, and clone-to-comments edges are intentionally
            // opaque in the generic profile. Record only these exact, verified
            // relationships so the package guard can distinguish them from an
            // unmodeled graph change.
            addedRelationshipIds.Add($"{PartPath(clonePart)}\0{leaf.LayoutRelationshipId}");
            changedParts.Add(PartPath(clonePart));
            changedParts.Add(RelationshipPartPath(clonePart));
        }
        changedParts.Add(PartPath(presentationPart));
        changedParts.Add(RelationshipPartPath(presentationPart));
        changedParts.Add("[Content_Types].xml");
    }

    // The first clone profile preserves the origin SlidePart byte-for-byte.
    // Allowing its removal in the same transaction would turn this leaf clone
    // into a topology-replacement primitive before the new graph has crossed
    // an export/reimport boundary, so reject it before any package mutation.
    private static void AssertCloneOriginsRetained(IReadOnlyList<PptxTargetSlideEntry> targets)
    {
        var retainedSources = targets
            .Where(target => !target.IsClone)
            .Select(target => target.Source.Part)
            .ToHashSet();
        foreach (var clone in targets.Where(target => target.IsClone))
        {
            if (!retainedSources.Contains(clone.Source.Part))
                throw new CodecException(
                    "unsupported_presentation_slide_clone",
                    $"Presentation clone {clone.TargetIndex + 1} cannot remove its origin slide in the same source-preserving export. Export and import the unchanged clone before changing source topology.",
                    PartPath(clone.Source.Part));
        }
    }

    private static PptxCloneLeaf AssertSourceSlideCanBeCloned(
        PresentationPart presentationPart,
        PptxTargetSlideEntry target,
        IReadOnlyDictionary<string, string> layoutIdByPartPath,
        IReadOnlyDictionary<string, string> slideIdByPartPath,
        IReadOnlySet<SlidePart> retainedSlideParts,
        PptxAssetCatalog assetCatalog)
    {
        var source = target.Source;
        var childParts = source.Part.Parts.ToArray();
        var layoutRelationshipCount = childParts.Count(pair => pair.OpenXmlPart is SlideLayoutPart);
        var imageEdges = childParts
            .Where(pair => pair.OpenXmlPart is ImagePart)
            .Select(pair => new PptxCloneImageEdge((ImagePart)pair.OpenXmlPart, pair.RelationshipId))
            .ToArray();
        var imageRelationshipIds = imageEdges
            .Select(edge => edge.RelationshipId)
            .ToHashSet(StringComparer.Ordinal);
        var externalHyperlinkEdges = source.Part.HyperlinkRelationships
            .Select(relationship => new PptxCloneExternalHyperlinkEdge(relationship.Uri, relationship.Id))
            .ToArray();
        var slideHyperlinkEdges = childParts
            .Where(pair => pair.OpenXmlPart is SlidePart)
            .Select(pair => new PptxCloneSlideHyperlinkEdge((SlidePart)pair.OpenXmlPart, pair.RelationshipId))
            .ToArray();
        var hyperlinkRelationshipIds = externalHyperlinkEdges
            .Select(edge => edge.RelationshipId)
            .Concat(slideHyperlinkEdges.Select(edge => edge.RelationshipId))
            .ToHashSet(StringComparer.Ordinal);
        var notesRelationshipCount = childParts.Count(pair => pair.OpenXmlPart is NotesSlidePart);
        var commentsRelationshipCount = childParts.Count(pair => pair.OpenXmlPart is SlideCommentsPart);
        var unsafeChildren = childParts
            .Where(pair => pair.OpenXmlPart is not SlideLayoutPart && pair.OpenXmlPart is not ImagePart && pair.OpenXmlPart is not SlidePart && pair.OpenXmlPart is not NotesSlidePart && pair.OpenXmlPart is not SlideCommentsPart)
            .Select(pair => pair.OpenXmlPart.RelationshipType)
            .Distinct(StringComparer.Ordinal)
            .Take(4)
            .ToArray();
        if (layoutRelationshipCount != 1 ||
            notesRelationshipCount > 1 ||
            commentsRelationshipCount > 1 ||
            unsafeChildren.Length > 0 ||
            source.Part.ExternalRelationships.Any() ||
            source.Part.DataPartReferenceRelationships.Any() ||
            hyperlinkRelationshipIds.Count > MaxCloneHyperlinkRelationships ||
            slideHyperlinkEdges.Any(edge => !retainedSlideParts.Contains(edge.Part)))
            throw UnsupportedSourceSlideClone(source, "it owns data parts, an unretained slide jump, too many hyperlinks, or another non-layout/image/run-hyperlink/notes/comments relationship");
        var root = source.Part.Slide ??
            throw new CodecException("missing_slide_root", $"Presentation source slide {source.Index + 1} has no slide root.", PartPath(source.Part));
        var common = root.CommonSlideData ??
            throw new CodecException("missing_common_slide_data", $"Presentation source slide {source.Index + 1} has no common slide data.", PartPath(source.Part));
        var tree = common.ShapeTree ??
            throw new CodecException("missing_shape_tree", $"Presentation source slide {source.Index + 1} has no shape tree.", PartPath(source.Part));
        if (root.Descendants<A.HyperlinkOnClick>().Take(MaxCloneHyperlinkRelationships + 1).Count() > MaxCloneHyperlinkRelationships)
            throw UnsupportedSourceSlideClone(source, $"it exceeds the {MaxCloneHyperlinkRelationships}-run-hyperlink markup budget");
        var sourceElements = ShapeElements(tree);
        var context = new PptxPartContext(source.Part, slideIdByPartPath, assets: assetCatalog);
        var elementIdsByNativeId = NativeElementIds(sourceElements, $"presentation/slide/{source.Index + 1}");
        var embeddedImageRelationshipIds = new HashSet<string>(StringComparer.Ordinal);
        var usedHyperlinkRelationshipIds = new HashSet<string>(StringComparer.Ordinal);
        for (var elementIndex = 0; elementIndex < sourceElements.Length; elementIndex++)
        {
            if (!TryCollectCanonicalCloneElement(
                    sourceElements[elementIndex],
                    $"presentation/slide/{source.Index + 1}",
                    elementIndex,
                    context,
                    elementIdsByNativeId,
                    imageRelationshipIds,
                    embeddedImageRelationshipIds,
                    usedHyperlinkRelationshipIds))
                throw UnsupportedSourceSlideClone(source, "its elements require a broader graph clone than the bounded shape-with-canonical-run-hyperlinks/inline-table/embedded-image/connector/recursive-group profile");
        }
        if (!imageRelationshipIds.SetEquals(embeddedImageRelationshipIds))
            throw UnsupportedSourceSlideClone(source, "it has an image relationship that is not exactly bound by a canonical embedded picture");
        if (!hyperlinkRelationshipIds.SetEquals(usedHyperlinkRelationshipIds))
            throw UnsupportedSourceSlideClone(source, "its hyperlink relationships are not exactly bound by canonical inline run click actions");
        var notes = AssertSourceSpeakerNotesCanBeCloned(source, target.Target);
        var comments = AssertSourceLegacyCommentsCanBeCloned(presentationPart, source, target.Target);
        if (target.Target.Name != (common.Name?.Value ?? string.Empty))
            throw UnsupportedSourceSlideClone(source, "the requested clone changes its source name");
        var layoutPart = source.Part.SlideLayoutPart ??
            throw UnsupportedSourceSlideClone(source, "it does not have a resolvable layout relationship");
        var expectedLayoutId = layoutIdByPartPath.GetValueOrDefault(PartPath(layoutPart));
        if (string.IsNullOrWhiteSpace(expectedLayoutId) || target.Target.LayoutId != expectedLayoutId)
            throw UnsupportedSourceSlideClone(source, "the requested clone changes its source layout binding");
        var sourceBinding = target.Target.CloneSource ??
            throw new CodecException("missing_presentation_slide_clone_binding", $"Presentation clone {target.TargetIndex + 1} is missing clone_source.", PartPath(source.Part));
        if (sourceBinding.LayoutRelationshipId != source.Part.GetIdOfPart(layoutPart) ||
            !sourceBinding.BackgroundSemanticSha256.Equals(BackgroundSemanticHash(PptxBackgroundCodec.Read(common)), StringComparison.OrdinalIgnoreCase))
            throw new CodecException("presentation_slide_clone_binding_mismatch", $"Presentation clone {target.TargetIndex + 1} does not match its source layout/background binding.", PartPath(source.Part));
        if (!BackgroundSemanticHash(target.Target.Background).Equals(BackgroundSemanticHash(PptxBackgroundCodec.Read(common)), StringComparison.OrdinalIgnoreCase))
            throw UnsupportedSourceSlideClone(source, "the requested clone changes its source background");
        if (sourceElements.Length != target.Target.Elements.Count)
            throw UnsupportedSourceSlideClone(source, "the requested clone changes source element topology");
        for (var elementIndex = 0; elementIndex < sourceElements.Length; elementIndex++)
        {
            var requested = target.Target.Elements[elementIndex];
            var binding = requested.Source ??
                throw new CodecException("missing_presentation_element_binding", $"Presentation clone {target.TargetIndex + 1} element {elementIndex + 1} is missing its source binding.", PartPath(source.Part));
            var original = ReadElement(sourceElements[elementIndex], source.Index, elementIndex, context, elementIdsByNativeId: elementIdsByNativeId);
            if (binding.ShapeTreeIndex != elementIndex ||
                !binding.ElementSha256.Equals(HashElement(sourceElements[elementIndex]), StringComparison.OrdinalIgnoreCase) ||
                binding.Editable != original.Source?.Editable ||
                binding.TextEditable != original.Source?.TextEditable ||
                !binding.SemanticSha256.Equals(original.Source?.SemanticSha256 ?? string.Empty, StringComparison.OrdinalIgnoreCase) ||
                !SemanticHash(original).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase) ||
                !SemanticHash(requested).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_slide_clone_mismatch", $"Presentation clone {target.TargetIndex + 1} element {elementIndex + 1} is not an unchanged source element.", PartPath(source.Part));
        }
        return new PptxCloneLeaf(
            layoutPart,
            source.Part.GetIdOfPart(layoutPart),
            imageEdges,
            externalHyperlinkEdges,
            slideHyperlinkEdges,
            notes,
            comments);
    }

    // A p:grpSp is a container, not a clone permission by itself. Its
    // descendants must recursively stay in the same canonical leaf profile,
    // and every nested picture must consume one verified SlidePart image
    // relationship. The clone operation still copies the original Slide XML;
    // this check only proves that no hidden graph edge is being smuggled in.
    private static bool TryCollectCanonicalCloneElement(
        OpenXmlElement element,
        string ownerId,
        int elementIndex,
        PptxPartContext context,
        IReadOnlyDictionary<uint, string> elementIdsByNativeId,
        ISet<string> imageRelationshipIds,
        ISet<string> embeddedImageRelationshipIds,
        ISet<string> usedHyperlinkRelationshipIds)
    {
        if (element is P.Shape shape)
            return IsSimpleShape(shape) && TryCollectCanonicalRunHyperlinks(shape, context, usedHyperlinkRelationshipIds);
        // A table is the one GraphicFrame profile that is entirely inline in
        // its containing XML. PptxTableCodec rejects charts and every other
        // GraphicFrame payload.
        if (element is P.GraphicFrame table)
            return !HasHyperlinkMarkup(table) && PptxTableCodec.TryRead(table, out _);
        if (element is P.Picture picture)
        {
            if (HasHyperlinkMarkup(picture)) return false;
            if (!PptxPictureCodec.TryRead(picture, context, out _)) return false;
            var relationshipId = picture.BlipFill?.GetFirstChild<A.Blip>()?.Embed?.Value ?? string.Empty;
            if (relationshipId.Length == 0 || !imageRelationshipIds.Contains(relationshipId)) return false;
            embeddedImageRelationshipIds.Add(relationshipId);
            return true;
        }
        // A bounded connection shape is entirely inline in its owning shape
        // tree. Its endpoints resolve through non-visual IDs in that same
        // cloned SlidePart, so it adds no OPC edge beyond the already-proved
        // layout/image/notes/comments graph.
        if (element is P.ConnectionShape connector)
            return !HasHyperlinkMarkup(connector) && TryReadConnector(connector, elementIdsByNativeId, out _);
        if (element is not P.GroupShape group) return false;

        var groupId = $"{ownerId}/element/{elementIndex + 1}";
        if (group.NonVisualGroupShapeProperties?.Descendants<A.HyperlinkOnClick>().Any() == true ||
            group.NonVisualGroupShapeProperties?.Descendants<A.HyperlinkOnHover>().Any() == true ||
            group.GroupShapeProperties?.Descendants<A.HyperlinkOnClick>().Any() == true ||
            group.GroupShapeProperties?.Descendants<A.HyperlinkOnHover>().Any() == true)
            return false;
        if (!TryReadGroup(group, groupId, context, elementIdsByNativeId, out _)) return false;
        var children = GroupElements(group);
        for (var childIndex = 0; childIndex < children.Length; childIndex++)
        {
            if (!TryCollectCanonicalCloneElement(
                    children[childIndex],
                    groupId,
                    childIndex,
                    context,
                    elementIdsByNativeId,
                    imageRelationshipIds,
                    embeddedImageRelationshipIds,
                    usedHyperlinkRelationshipIds))
                return false;
        }
        return true;
    }

    private static bool TryCollectCanonicalRunHyperlinks(
        P.Shape shape,
        PptxPartContext context,
        ISet<string> usedRelationshipIds)
    {
        if (shape.Descendants<A.HyperlinkOnHover>().Any()) return false;
        var clicks = shape.Descendants<A.HyperlinkOnClick>().ToArray();
        if (clicks.Length > MaxCloneHyperlinkRelationships) return false;
        foreach (var click in clicks)
        {
            if (click.Parent is not A.RunProperties properties ||
                properties.Parent is not (A.Run or A.Break or A.Field) ||
                properties.Elements<A.HyperlinkOnClick>().Count() != 1 ||
                !PptxHyperlinkCodec.TryRead(click, context, out _))
                return false;
            var relationshipId = click.Id?.Value ?? string.Empty;
            if (relationshipId.Length > 0) usedRelationshipIds.Add(relationshipId);
        }
        return true;
    }

    private static bool HasHyperlinkMarkup(OpenXmlElement element) =>
        element.Descendants<A.HyperlinkOnClick>().Any() || element.Descendants<A.HyperlinkOnHover>().Any();

    private static PptxCloneNotesLeaf? AssertSourceSpeakerNotesCanBeCloned(PptxSourceSlideEntry source, PresentationSlide target)
    {
        var notesPart = source.Part.NotesSlidePart;
        if (notesPart is null)
        {
            if (target.SpeakerNotes is not null)
                throw UnsupportedSourceSlideClone(source, "the requested clone adds speaker notes where the origin has no notes part");
            return null;
        }
        var original = PptxSpeakerNotesCodec.Read(source.Part) ??
            throw UnsupportedSourceSlideClone(source, "its notes relationship has no readable notes root");
        var requested = target.SpeakerNotes;
        if (requested?.Source is not { } binding)
            throw new CodecException("missing_presentation_notes_binding", $"Presentation clone {source.Index + 1} speaker notes are missing their source binding.", PartPath(notesPart));
        var sourceBinding = original.Source;
        if (!requested.Text.Equals(original.Text, StringComparison.Ordinal) ||
            !binding.PartPath.Equals(sourceBinding.PartPath, StringComparison.OrdinalIgnoreCase) ||
            !binding.RelationshipId.Equals(sourceBinding.RelationshipId, StringComparison.Ordinal) ||
            !binding.NotesXmlSha256.Equals(sourceBinding.NotesXmlSha256, StringComparison.OrdinalIgnoreCase) ||
            !binding.SemanticSha256.Equals(sourceBinding.SemanticSha256, StringComparison.OrdinalIgnoreCase) ||
            binding.Editable != sourceBinding.Editable)
            throw new CodecException("presentation_slide_clone_mismatch", $"Presentation clone {source.Index + 1} speaker notes are not unchanged source notes.", PartPath(notesPart));

        var relationships = notesPart.Parts.ToArray();
        var masterEdges = relationships.Where(pair => pair.OpenXmlPart is NotesMasterPart).ToArray();
        var backReferenceEdges = relationships.Where(pair => ReferenceEquals(pair.OpenXmlPart, source.Part)).ToArray();
        var unsafeChildren = relationships
            .Where(pair => pair.OpenXmlPart is not NotesMasterPart && !ReferenceEquals(pair.OpenXmlPart, source.Part))
            .Select(pair => pair.OpenXmlPart.RelationshipType)
            .Distinct(StringComparer.Ordinal)
            .Take(4)
            .ToArray();
        if (masterEdges.Length != 1 || backReferenceEdges.Length != 1 || unsafeChildren.Length > 0 ||
            notesPart.ExternalRelationships.Any() || notesPart.HyperlinkRelationships.Any() || notesPart.DataPartReferenceRelationships.Any())
            throw UnsupportedSourceSlideClone(source, "its notes part has relationships beyond one NotesMaster and one back-reference to the source slide");
        var masterPart = (NotesMasterPart)masterEdges[0].OpenXmlPart;
        return new PptxCloneNotesLeaf(
            notesPart,
            source.Part.GetIdOfPart(notesPart),
            masterPart,
            masterEdges[0].RelationshipId,
            backReferenceEdges[0].RelationshipId);
    }

    private static PptxCloneLegacyCommentsLeaf? AssertSourceLegacyCommentsCanBeCloned(
        PresentationPart presentationPart,
        PptxSourceSlideEntry source,
        PresentationSlide target)
    {
        var commentsPart = source.Part.SlideCommentsPart;
        if (commentsPart is null)
        {
            if (target.LegacyComments.Count > 0)
                throw UnsupportedSourceSlideClone(source, "the requested clone adds legacy comments where the origin has no comments part");
            return null;
        }
        var profile = PptxLegacyCommentsCodec.Profile(presentationPart, source.Part, source.Index);
        if (!profile.Supported)
            throw UnsupportedSourceSlideClone(source, "its legacy comments are outside the bounded single-text profile");
        if (!PptxLegacyCommentsCodec.Equivalent(profile.Comments, target.LegacyComments))
            throw new CodecException(
                "presentation_slide_clone_mismatch",
                $"Presentation clone {source.Index + 1} legacy comments are not unchanged source comments.",
                PartPath(commentsPart));
        if (commentsPart.Parts.Any() || commentsPart.ExternalRelationships.Any() ||
            commentsPart.HyperlinkRelationships.Any() || commentsPart.DataPartReferenceRelationships.Any())
            throw UnsupportedSourceSlideClone(source, "its comments part has child, external, hyperlink, or data relationships");

        // A legacy p:cm refers to authorId in the presentation-wide
        // CommentAuthorsPart rather than through a relationship on the
        // comments part. The source profile proves every ID resolves; this
        // additional graph check makes the shared catalog immutable and
        // explicit in the clone plan.
        var authorsPart = presentationPart.CommentAuthorsPart;
        var authorEdges = presentationPart.Parts
            .Where(pair => ReferenceEquals(pair.OpenXmlPart, authorsPart))
            .ToArray();
        if (authorsPart is null || authorEdges.Length != 1 || authorsPart.Parts.Any() ||
            authorsPart.ExternalRelationships.Any() || authorsPart.HyperlinkRelationships.Any() ||
            authorsPart.DataPartReferenceRelationships.Any())
            throw UnsupportedSourceSlideClone(source, "its legacy comment author catalog is missing or has an unbounded relationship graph");
        return new PptxCloneLegacyCommentsLeaf(
            commentsPart,
            source.Part.GetIdOfPart(commentsPart),
            authorsPart,
            authorEdges[0].RelationshipId);
    }

    private static CodecException UnsupportedSourceSlideClone(PptxSourceSlideEntry source, string reason) =>
        new(
            "unsupported_presentation_slide_clone",
            $"Source-preserving PPTX cloning is limited to an unchanged shape-with-canonical-run-hyperlinks/inline-table/embedded-image/bounded-connector/recursive-group layout leaf with optional closed notes and legacy-comment leaves; slide {source.Index + 1} cannot be cloned because {reason}. Use an explicit broader OPC graph-clone operation for this package.",
            PartPath(source.Part));

    private static bool ReorderSourceSlideIdList(PresentationPart presentationPart, IReadOnlyList<PptxTargetSlideEntry> targets)
    {
        var root = presentationPart.Presentation ??
            throw new CodecException("missing_presentation_root", "PPTX package has no Presentation root.", "ppt/presentation.xml");
        var list = root.SlideIdList ??
            throw new CodecException("missing_slide_id_list", "PPTX package has no slide ID list.", "ppt/presentation.xml");
        var sourceIds = list.Elements<P.SlideId>().ToArray();
        var requestedRelationshipIds = targets.Select(target => target.OutputSlideId.RelationshipId?.Value ?? string.Empty).ToArray();
        if (sourceIds.Select(slideId => slideId.RelationshipId?.Value ?? string.Empty).SequenceEqual(requestedRelationshipIds, StringComparer.Ordinal)) return false;

        // Keep any future extension child in its original relative position;
        // only replace the ordered p:sldId entries with exact source clones.
        var firstNonSlideId = list.ChildElements.FirstOrDefault(item => item is not P.SlideId);
        var reordered = targets.Select(target => (P.SlideId)target.OutputSlideId.CloneNode(true)).ToArray();
        foreach (var sourceId in sourceIds) sourceId.Remove();
        foreach (var slideId in reordered)
        {
            if (firstNonSlideId is null) list.Append(slideId);
            else list.InsertBefore(slideId, firstNonSlideId);
        }
        root.Save();
        return true;
    }

    private static string MasterTextStylesSemanticHash(PresentationMasterTextStyles? source)
    {
        var semantic = source?.Clone() ?? new PresentationMasterTextStyles();
        PptxMasterTextStylesCodec.NormalizeSemantics(semantic);
        return Hash(semantic.ToByteArray());
    }

    private static string BackgroundSemanticHash(PresentationBackground? source) =>
        Hash((source ?? new PresentationBackground()).ToByteArray());

    private static bool ApplyPlaceholders(
        P.ShapeTree shapeTree,
        string ownerId,
        IList<PresentationPlaceholder> requested,
        PptxPartContext partContext,
        string partPath)
    {
        var originals = PptxPlaceholderCodec.Read(shapeTree, ownerId, partContext);
        if (originals.Count != requested.Count)
            throw new CodecException(
                "presentation_placeholder_topology_changed",
                $"Source-preserving PPTX export requires {ownerId}'s original {originals.Count}-placeholder topology; the artifact contains {requested.Count} placeholders.",
                partPath);
        var changed = false;
        for (var index = 0; index < originals.Count; index++)
        {
            var original = originals[index];
            var target = requested[index];
            var sourceBinding = original.Source!;
            var binding = target.Source ?? throw new CodecException(
                "missing_presentation_placeholder_binding",
                $"Presentation placeholder {index + 1} under {ownerId} is missing its source binding.",
                partPath);
            var sourceShape = PptxPlaceholderCodec.BoundShape(shapeTree, original);
            if (sourceShape is null || target.Id != original.Id ||
                binding.ShapeTreeIndex != sourceBinding.ShapeTreeIndex ||
                !binding.ElementSha256.Equals(sourceBinding.ElementSha256, StringComparison.OrdinalIgnoreCase) ||
                !binding.SemanticSha256.Equals(sourceBinding.SemanticSha256, StringComparison.OrdinalIgnoreCase) ||
                binding.Editable != sourceBinding.Editable ||
                binding.DirectFramePresenceEditable != sourceBinding.DirectFramePresenceEditable ||
                binding.TextEditable != sourceBinding.TextEditable ||
                !binding.ElementSha256.Equals(PptxPlaceholderCodec.ElementHash(sourceShape), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_placeholder_binding_mismatch",
                    $"Presentation placeholder {index + 1} under {ownerId} does not match its hash-bound source element.",
                    partPath);
            if (!PptxPlaceholderCodec.SemanticHash(original).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_placeholder_source_semantics_mismatch",
                    $"Presentation placeholder {index + 1} under {ownerId} does not match its source semantic binding.",
                    partPath);
            if (PptxPlaceholderCodec.SemanticHash(target).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase)) continue;
            if (!binding.Editable)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation placeholder {index + 1} under {ownerId} has no safely editable semantic component in this codec slice.",
                    partPath);
            PptxPlaceholderCodec.Apply(sourceShape, original, target, partContext);
            changed = true;
        }
        return changed;
    }

    private static IEnumerable<PresentationTextParagraph> MasterStyleParagraphs(PresentationMasterTextStyles? source) =>
        source is null
            ? []
            : source.TitleLevels.Concat(source.BodyLevels).Concat(source.OtherLevels);

    private static void TrackContextChanges(
        OpenXmlPart owner,
        PptxPartContext context,
        ISet<string> changedParts,
        ISet<string> addedRelationshipIds,
        ISet<string> addedPartPaths)
    {
        if (context.RelationshipsChanged)
        {
            changedParts.Add(RelationshipPartPath(owner));
            foreach (var id in context.AddedRelationshipIds)
                addedRelationshipIds.Add($"{PartPath(owner)}\0{id}");
        }
        foreach (var path in context.AddedPartPaths)
        {
            changedParts.Add(path);
            addedPartPaths.Add(path);
            changedParts.Add("[Content_Types].xml");
        }
    }

    private static string ShapeResidualHash(P.Shape source, PptxPartContext slideContext)
    {
        var shape = (P.Shape)source.CloneNode(true);
        if (shape.NonVisualShapeProperties?.NonVisualDrawingProperties is { } nonVisual) nonVisual.Name = string.Empty;
        if (shape.NonVisualShapeProperties?.NonVisualShapeDrawingProperties is { } drawingProperties) drawingProperties.TextBox = null;
        if (shape.ShapeProperties is { } properties)
        {
            if (properties.Transform2D is { } transform)
            {
                if (transform.Offset is { } offset) { offset.X = 0L; offset.Y = 0L; }
                if (transform.Extents is { } extents) { extents.Cx = 1L; extents.Cy = 1L; }
                PptxShapeTransformCodec.Scrub(transform);
            }
            properties.GetFirstChild<A.CustomGeometry>()?.Remove();
            if (properties.GetFirstChild<A.PresetGeometry>() is { } geometry) geometry.Preset = A.ShapeTypeValues.Rectangle;
            else properties.InsertAfter(new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }, properties.GetFirstChild<A.Transform2D>());
            properties.GetFirstChild<A.EffectList>()?.Remove();
            foreach (var fill in properties.ChildElements.Where(child => child is A.NoFill or A.SolidFill).ToArray()) fill.Remove();
            if (properties.GetFirstChild<A.Outline>() is { } outline)
            {
                outline.Width = 0;
                foreach (var fill in outline.ChildElements.Where(child => child is A.NoFill or A.SolidFill).ToArray()) fill.Remove();
            }
        }
        PptxTextCodec.ScrubModeledContent(shape.TextBody, slideContext);
        return HashElement(shape);
    }

    private static string ConnectorResidualHash(P.ConnectionShape source)
    {
        var connector = (P.ConnectionShape)source.CloneNode(true);
        if (connector.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties is { } nonVisual) nonVisual.Name = string.Empty;
        if (connector.NonVisualConnectionShapeProperties?.NonVisualConnectorShapeDrawingProperties is { } drawingProperties)
            drawingProperties.RemoveAllChildren();
        connector.ShapeProperties?.RemoveAllChildren();
        return HashElement(connector);
    }

    private static string ChartFrameResidualHash(P.GraphicFrame source)
    {
        var chart = (P.GraphicFrame)source.CloneNode(true);
        PptxChartCodec.ScrubFrame(chart);
        return HashElement(chart);
    }

    private static string PictureResidualHash(P.Picture source)
    {
        var picture = (P.Picture)source.CloneNode(true);
        PptxPictureCodec.ScrubModeledContent(picture);
        return HashElement(picture);
    }

    private static string TableResidualHash(P.GraphicFrame source)
    {
        var table = (P.GraphicFrame)source.CloneNode(true);
        PptxTableCodec.ScrubModeledContent(table);
        return HashElement(table);
    }

    private static string NativeObjectResidualHash(OpenXmlElement source)
    {
        var clone = source.CloneNode(true);
        if (clone.Descendants<P.NonVisualDrawingProperties>().FirstOrDefault() is { } nonVisual)
            nonVisual.Name = string.Empty;
        if (clone is P.GraphicFrame graphicFrame && graphicFrame.Transform is { } transform)
        {
            ScrubFrame(transform);
            if (PptxNativeObjectCatalog.Classify(clone) == "oleObject" && graphicFrame.Descendants<A.Transform2D>().FirstOrDefault() is { } preview)
                ScrubFrame(preview);
        }
        else if (clone is P.GroupShape group && group.GetFirstChild<P.GroupShapeProperties>()?.GetFirstChild<A.TransformGroup>() is { } groupTransform)
        {
            ScrubFrame(groupTransform);
        }
        return HashElement(clone);
    }

    private static void ScrubFrame(P.Transform transform)
    {
        transform.Offset!.X = 0L;
        transform.Offset.Y = 0L;
        transform.Extents!.Cx = 1L;
        transform.Extents.Cy = 1L;
    }

    private static void ScrubFrame(A.Transform2D transform)
    {
        if (transform.Offset is { } offset) { offset.X = 0L; offset.Y = 0L; }
        if (transform.Extents is { } extents) { extents.Cx = 1L; extents.Cy = 1L; }
    }

    private static void ScrubFrame(A.TransformGroup transform)
    {
        transform.Offset!.X = 0L;
        transform.Offset.Y = 0L;
        transform.Extents!.Cx = 1L;
        transform.Extents.Cy = 1L;
    }

    private static string MasterResidualHash(P.SlideMaster source, PptxPartContext partContext)
    {
        var master = (P.SlideMaster)source.CloneNode(true);
        PptxMasterTextStylesCodec.ScrubModeledContent(master, partContext);
        PptxBackgroundCodec.ScrubModeledContent(master.CommonSlideData);
        PptxPlaceholderCodec.ScrubModeledContent(master.CommonSlideData?.ShapeTree, partContext);
        return HashElement(master);
    }

    private static string LayoutResidualHash(P.SlideLayout source, PptxPartContext partContext)
    {
        var layout = (P.SlideLayout)source.CloneNode(true);
        PptxBackgroundCodec.ScrubModeledContent(layout.CommonSlideData);
        PptxPlaceholderCodec.ScrubModeledContent(layout.CommonSlideData?.ShapeTree, partContext);
        return HashElement(layout);
    }

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute => allowed.Contains(attribute.LocalName));
    }

    private static Dictionary<string, string> PackagePartHashes(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        return archive.Entries.Where(entry => !entry.FullName.EndsWith('/')).ToDictionary(
            entry => entry.FullName,
            entry =>
            {
                using var source = entry.Open();
                using var copy = new MemoryStream();
                source.CopyTo(copy);
                return Hash(copy.ToArray());
            },
            StringComparer.OrdinalIgnoreCase);
    }

    private static void ValidateOutputBudget(byte[] bytes, EffectiveCodecLimits limits)
    {
        if ((ulong)bytes.LongLength > limits.MaxInputBytes)
            throw new CodecException("output_budget_exceeded", $"Generated PPTX has {bytes.LongLength} bytes and exceeds max_input_bytes ({limits.MaxInputBytes}).");
    }

    private const int MaxOffice2021ValidationErrors = 256;

    private static void ValidateOffice2021(byte[] bytes)
    {
        var errors = Office2021ValidationErrors(bytes);
        if (errors.Length == 0) return;
        var detail = string.Join("; ", errors.Take(8).Select(error => error.Detail));
        throw new CodecException("openxml_validation_failed", $"Generated PPTX is not valid Office 2021 Open XML: {detail}");
    }

    // An imported package may use valid-in-host extension markup that the
    // Open XML SDK's Office 2021 validator does not recognize. Source-bound
    // export may retain such diagnostics, but must never introduce a new one.
    // This comparison is intentionally unavailable to source-free export.
    private static int ValidateOffice2021AgainstSource(byte[] sourceBytes, byte[] outputBytes)
    {
        var sourceErrors = Office2021ValidationErrors(sourceBytes);
        var sourceSignatures = sourceErrors
            .Select(error => error.Signature)
            .ToHashSet(StringComparer.Ordinal);
        var introduced = Office2021ValidationErrors(outputBytes)
            .Where(error => !sourceSignatures.Contains(error.Signature))
            .Take(8)
            .ToArray();
        if (introduced.Length == 0) return sourceErrors.Length;
        var detail = string.Join("; ", introduced.Select(error => error.Detail));
        throw new CodecException("openxml_validation_failed", $"Source-preserving PPTX export introduced Office 2021 Open XML validation error(s): {detail}");
    }

    private sealed record Office2021ValidationError(string Signature, string Detail);

    private static Office2021ValidationError[] Office2021ValidationErrors(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var package = PresentationDocument.Open(stream, isEditable: false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2021)
            .Validate(package)
            .Take(MaxOffice2021ValidationErrors + 1)
            .Select(error =>
            {
                var part = error.Part?.Uri.ToString() ?? "package";
                var path = error.Path?.XPath ?? string.Empty;
                var detail = $"{(path.Length == 0 ? part : path)}: {error.Description}";
                var signature = string.Join("\u001f", part, path, error.Id ?? string.Empty, error.Description ?? string.Empty);
                return new Office2021ValidationError(signature, detail);
            })
            .ToArray();
        if (errors.Length <= MaxOffice2021ValidationErrors) return errors;
        throw new CodecException(
            "openxml_validation_budget_exceeded",
            $"PPTX validation produced more than {MaxOffice2021ValidationErrors} errors; refusing an unbounded validation result.");
    }

}
