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

namespace OpenOffice.OpenXmlCodec;

internal sealed record PptxImportResult(ArtifactEnvelope Artifact, IReadOnlyList<Diagnostic> Diagnostics);
internal sealed record PptxExportResult(byte[] File, IReadOnlyList<Diagnostic> Diagnostics);

internal static class PptxCodec
{
    private const long DefaultSlideWidthEmu = 12_192_000;
    private const long DefaultSlideHeightEmu = 6_858_000;

    internal static PptxImportResult Import(byte[] bytes, EffectiveCodecLimits limits)
    {
        var opaque = PackageGuards.ValidateAndCollectOpaque(bytes, limits, OpcPackageProfile.Pptx);
        var diagnostics = new List<Diagnostic>();
        var opaqueCount = opaque.Parts.Count + opaque.PackageRelationships.Count;
        if (opaqueCount > 0)
            diagnostics.Add(CodecProtocol.Warning(
                "opaque_content_retained",
                $"Retained {opaqueCount} unsupported OPC parts or relationships with a hash-bound source package snapshot for loss-aware export.",
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

        var artifact = new PresentationArtifact
        {
            Id = "presentation/1",
            Name = "Imported presentation",
            SlideWidthEmu = presentationRoot.SlideSize?.Cx?.Value ?? DefaultSlideWidthEmu,
            SlideHeightEmu = presentationRoot.SlideSize?.Cy?.Value ?? DefaultSlideHeightEmu,
        };
        ulong semanticItems = 0;
        for (var slideIndex = 0; slideIndex < slideIds.Length; slideIndex++)
        {
            var slideId = slideIds[slideIndex];
            var relationshipId = slideId.RelationshipId?.Value ?? string.Empty;
            var slidePart = slideParts[slideIndex];
            var slideRoot = slidePart.Slide ??
                throw new CodecException("missing_slide_root", $"Presentation slide {slideIndex + 1} has no slide root.", PartPath(slidePart));
            var shapeTree = slideRoot.CommonSlideData?.ShapeTree ??
                throw new CodecException("missing_shape_tree", $"Presentation slide {slideIndex + 1} has no shape tree.", PartPath(slidePart));
            var elements = ShapeElements(shapeTree);
            var target = new PresentationSlide
            {
                Id = $"presentation/slide/{slideIndex + 1}",
                Name = slideRoot.CommonSlideData?.Name?.Value ?? $"Slide {slideIndex + 1}",
                Source = new PresentationSlideSourceBinding
                {
                    SlideIndex = checked((uint)slideIndex),
                    PartPath = PartPath(slidePart),
                    RelationshipId = relationshipId,
                    SlideXmlSha256 = HashElement(slideRoot),
                },
            };
            var hyperlinkContext = new PptxHyperlinkContext(slidePart, slideIdByPartPath);
            for (var elementIndex = 0; elementIndex < elements.Length; elementIndex++)
            {
                semanticItems++;
                if (semanticItems > limits.MaxCells)
                    throw new CodecException("presentation_item_budget_exceeded", $"PPTX presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).", PartPath(slidePart));
                target.Elements.Add(ReadElement(elements[elementIndex], slideIndex, elementIndex, hyperlinkContext));
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
                Producer = "open-office-artifact-tool/OpenXmlWasm",
            },
        };
        envelope.Diagnostics.Add(diagnostics);
        return new PptxImportResult(envelope, diagnostics);
    }

    internal static PptxExportResult Export(ArtifactEnvelope envelope, EffectiveCodecLimits limits, bool allowLossy)
    {
        ValidateEnvelope(envelope, limits);
        var opaqueCount = (envelope.OpaqueOpc?.Parts.Count ?? 0) +
                          (envelope.OpaqueOpc?.PackageRelationships.Count ?? 0);
        if (envelope.OpaqueOpc?.SourcePackage is { Data.IsEmpty: false })
            return ExportPreservingSource(envelope, limits, opaqueCount);
        if (opaqueCount > 0 && !allowLossy)
            throw new CodecException(
                "opaque_content_requires_preservation",
                "Presentation contains opaque OPC parts or relationships but its validated source package snapshot is unavailable; pass allow_lossy only when discarding them is intentional.");

        var diagnostics = new List<Diagnostic>();
        if (opaqueCount > 0)
            diagnostics.Add(CodecProtocol.Warning(
                "opaque_content_discarded",
                $"Discarded {opaqueCount} opaque OPC parts or relationships under explicit allow_lossy policy."));

        using var stream = new MemoryStream();
        using (var package = PresentationDocument.Create(stream, PresentationDocumentType.Presentation, autoSave: true))
            BuildPresentation(package, envelope.Presentation);
        var bytes = stream.ToArray();
        ValidateOutputBudget(bytes, limits);
        ValidateOffice2021(bytes);
        return new PptxExportResult(bytes, diagnostics);
    }

    private static PptxExportResult ExportPreservingSource(ArtifactEnvelope envelope, EffectiveCodecLimits limits, int opaqueCount)
    {
        var sourceBytes = PackageGuards.ValidateSourcePackage(envelope.OpaqueOpc, envelope.Source, limits, OpcPackageProfile.Pptx);
        using var stream = new MemoryStream();
        stream.Write(sourceBytes);
        stream.Position = 0;
        var changedParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using (var package = PresentationDocument.Open(stream, isEditable: true, new OpenSettings { AutoSave = false }))
        {
            var presentationPart = package.PresentationPart ??
                throw new CodecException("missing_presentation_part", "PPTX package has no Presentation part.", "ppt/presentation.xml");
            var slideIds = presentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>().ToArray() ?? [];
            if (slideIds.Length != envelope.Presentation.Slides.Count)
                throw new CodecException(
                    "presentation_topology_changed",
                    $"Source-preserving PPTX export requires the original {slideIds.Length}-slide topology; the artifact contains {envelope.Presentation.Slides.Count} slides.",
                    "ppt/presentation.xml");
            var slideParts = ResolveSlideParts(presentationPart, slideIds);
            var slideIdByPartPath = slideParts
                .Select((part, index) => (Path: PartPath(part), Id: envelope.Presentation.Slides[index].Id))
                .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
            var slidePartById = slideParts
                .Select((part, index) => (Part: part, Id: envelope.Presentation.Slides[index].Id))
                .ToDictionary(item => item.Id, item => item.Part, StringComparer.Ordinal);

            ulong semanticItems = 0;
            for (var slideIndex = 0; slideIndex < slideIds.Length; slideIndex++)
            {
                var sourceSlideId = slideIds[slideIndex];
                var relationshipId = sourceSlideId.RelationshipId?.Value ?? string.Empty;
                var slidePart = slideParts[slideIndex];
                var slideRoot = slidePart.Slide ??
                    throw new CodecException("missing_slide_root", $"Presentation slide {slideIndex + 1} has no slide root.", PartPath(slidePart));
                var target = envelope.Presentation.Slides[slideIndex];
                var binding = target.Source ?? throw new CodecException(
                    "missing_presentation_slide_binding",
                    $"Presentation slide {slideIndex + 1} is missing its source binding.",
                    "ppt/presentation.xml");
                if (binding.SlideIndex != slideIndex ||
                    !binding.PartPath.Equals(PartPath(slidePart), StringComparison.OrdinalIgnoreCase) ||
                    !binding.RelationshipId.Equals(relationshipId, StringComparison.Ordinal) ||
                    !binding.SlideXmlSha256.Equals(HashElement(slideRoot), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_slide_binding_mismatch",
                        $"Presentation slide {slideIndex + 1} does not match its hash-bound source slide.",
                        PartPath(slidePart));

                var shapeTree = slideRoot.CommonSlideData?.ShapeTree ??
                    throw new CodecException("missing_shape_tree", $"Presentation slide {slideIndex + 1} has no shape tree.", PartPath(slidePart));
                var sourceElements = ShapeElements(shapeTree);
                if (sourceElements.Length != target.Elements.Count)
                    throw new CodecException(
                        "presentation_element_topology_changed",
                        $"Source-preserving PPTX export requires slide {slideIndex + 1}'s original {sourceElements.Length}-element topology; the artifact contains {target.Elements.Count} elements.",
                        PartPath(slidePart));

                var changed = false;
                var hyperlinkContext = new PptxHyperlinkContext(slidePart, slideIdByPartPath, slidePartById);
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
                    var original = ReadElement(sourceElement, slideIndex, elementIndex, hyperlinkContext);
                    if (!SemanticHash(original).Equals(elementBinding.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "presentation_source_semantics_mismatch",
                            $"Presentation slide {slideIndex + 1} element {elementIndex + 1} source semantics do not match its binding.",
                            PartPath(slidePart));
                    if (SemanticHash(requested).Equals(elementBinding.SemanticSha256, StringComparison.OrdinalIgnoreCase)) continue;
                    if (!elementBinding.Editable || sourceElement is not P.Shape sourceShape || requested.ContentCase != PresentationElement.ContentOneofCase.Shape)
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            $"Presentation slide {slideIndex + 1} element {elementIndex + 1} is preserved but not safely editable by this codec slice.",
                            PartPath(slidePart));
                    ApplyShape(sourceShape, requested, hyperlinkContext);
                    changed = true;
                }
                if (changed)
                {
                    slideRoot.Save();
                    changedParts.Add(PartPath(slidePart));
                }
                if (hyperlinkContext.RelationshipsChanged)
                    changedParts.Add(RelationshipPartPath(slidePart));
            }
        }

        var bytes = stream.ToArray();
        ValidateOutputBudget(bytes, limits);
        ValidateOffice2021(bytes);
        AssertPackagePartsUnchangedExcept(sourceBytes, bytes, changedParts);
        ValidatePreservedSlideElements(sourceBytes, bytes, envelope.Presentation);
        var outputOpaque = PackageGuards.ValidateAndCollectOpaque(bytes, limits, OpcPackageProfile.Pptx, includeSourcePackage: false);
        PackageGuards.AssertOpaqueGraphMatches(envelope.OpaqueOpc, outputOpaque, "opaque_content_not_preserved");
        var diagnostics = new List<Diagnostic>();
        if (opaqueCount > 0)
            diagnostics.Add(CodecProtocol.Warning(
                "opaque_content_preserved",
                $"Preserved {opaqueCount} opaque OPC parts or relationships while updating modeled presentation content."));
        return new PptxExportResult(bytes, diagnostics);
    }

    private static PresentationElement ReadElement(OpenXmlElement source, int slideIndex, int elementIndex, PptxHyperlinkContext hyperlinkContext)
    {
        var element = new PresentationElement
        {
            Id = $"presentation/slide/{slideIndex + 1}/element/{elementIndex + 1}",
            Name = ElementName(source, elementIndex),
        };
        var editable = source is P.Shape shape && IsSimpleShape(shape);
        if (source is P.Shape sourceShape)
            element.Shape = ReadShape(sourceShape, hyperlinkContext);
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
        }
        element.Source = new PresentationElementSourceBinding
        {
            ShapeTreeIndex = checked((uint)elementIndex),
            ElementSha256 = HashElement(source),
            Editable = editable,
        };
        element.Source.SemanticSha256 = SemanticHash(element);
        return element;
    }

    private static PresentationShape ReadShape(P.Shape shape, PptxHyperlinkContext hyperlinkContext)
    {
        var frame = ReadFrame(shape);
        var properties = shape.ShapeProperties;
        var textBody = PptxTextCodec.Read(shape.TextBody, hyperlinkContext);
        return new PresentationShape
        {
            Geometry = Geometry(properties),
            LeftEmu = frame.Left,
            TopEmu = frame.Top,
            WidthEmu = frame.Width,
            HeightEmu = frame.Height,
            Text = PptxTextCodec.Flatten(textBody),
            TextBody = textBody,
            FillRgb = PptxColor.SolidRgb(properties?.GetFirstChild<A.SolidFill>()),
            LineRgb = PptxColor.SolidRgb(properties?.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>()),
            LineWidthEmu = properties?.GetFirstChild<A.Outline>()?.Width?.Value ?? 0,
        };
    }

    private static bool IsSimpleShape(P.Shape shape)
    {
        if (shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>() is not null) return false;
        if (shape.ShapeStyle is not null) return false;
        var properties = shape.ShapeProperties;
        var transform = properties?.Transform2D;
        if (properties is null || transform?.Offset is null || transform.Extents is null || transform.Rotation is not null || transform.HorizontalFlip is not null || transform.VerticalFlip is not null) return false;
        if (Geometry(properties) is not ("rect" or "ellipse")) return false;
        if (!SimpleFill(properties)) return false;
        var outline = properties.GetFirstChild<A.Outline>();
        if (outline is not null && !SimpleFill(outline)) return false;
        if (properties.ChildElements.Any(child => child is not A.Transform2D and not A.PresetGeometry and not A.NoFill and not A.SolidFill and not A.Outline)) return false;
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

    private static void ApplyShape(P.Shape shape, PresentationElement source, PptxHyperlinkContext hyperlinkContext)
    {
        var semantic = source.Shape;
        var properties = shape.ShapeProperties ??= new P.ShapeProperties();
        var transform = properties.Transform2D ??= new A.Transform2D();
        var offset = transform.Offset ??= new A.Offset();
        offset.X = semantic.LeftEmu;
        offset.Y = semantic.TopEmu;
        var extents = transform.Extents ??= new A.Extents();
        extents.Cx = semantic.WidthEmu;
        extents.Cy = semantic.HeightEmu;
        var geometry = properties.GetFirstChild<A.PresetGeometry>();
        if (geometry is null)
        {
            geometry = new A.PresetGeometry(new A.AdjustValueList());
            properties.InsertAfter(geometry, transform);
        }
        geometry.Preset = semantic.Geometry == "ellipse" ? A.ShapeTypeValues.Ellipse : A.ShapeTypeValues.Rectangle;
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
        PptxTextCodec.Apply(shape, semantic, hyperlinkContext);
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

    private static void BuildPresentation(PresentationDocument package, PresentationArtifact artifact)
    {
        var presentationPart = package.AddPresentationPart();
        var masterPart = presentationPart.AddNewPart<SlideMasterPart>("rIdMaster1");
        var layoutPart = masterPart.AddNewPart<SlideLayoutPart>("rIdLayout1");
        var themePart = masterPart.AddNewPart<ThemePart>("rIdTheme1");
        layoutPart.AddPart(masterPart, "rIdMaster1");

        themePart.Theme = BasicTheme();
        layoutPart.SlideLayout = new P.SlideLayout(
            new P.CommonSlideData(BasicShapeTree()) { Name = "Blank" },
            new P.ColorMapOverride(new A.MasterColorMapping()))
        { Type = P.SlideLayoutValues.Blank, Preserve = true };
        masterPart.SlideMaster = new P.SlideMaster(
            new P.CommonSlideData(BasicShapeTree()) { Name = "Office Clean Room" },
            BasicColorMap(),
            new P.SlideLayoutIdList(new P.SlideLayoutId { Id = 2_147_483_649U, RelationshipId = "rIdLayout1" }),
            new P.TextStyles(new P.TitleStyle(), new P.BodyStyle(), new P.OtherStyle()));

        var slideIdList = new P.SlideIdList();
        var slideParts = new SlidePart[artifact.Slides.Count];
        for (var slideIndex = 0; slideIndex < artifact.Slides.Count; slideIndex++)
        {
            var source = artifact.Slides[slideIndex];
            var relationshipId = $"rIdSlide{slideIndex + 1}";
            var slidePart = presentationPart.AddNewPart<SlidePart>(relationshipId);
            slideParts[slideIndex] = slidePart;
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
        for (var slideIndex = 0; slideIndex < artifact.Slides.Count; slideIndex++)
        {
            var source = artifact.Slides[slideIndex];
            var slidePart = slideParts[slideIndex];
            var shapeTree = slidePart.Slide!.CommonSlideData!.ShapeTree!;
            var hyperlinkContext = new PptxHyperlinkContext(slidePart, slideIdByPartPath, slidePartById);
            uint nativeId = 2;
            foreach (var element in source.Elements)
            {
                if (element.ContentCase != PresentationElement.ContentOneofCase.Shape)
                    throw new CodecException("unsupported_presentation_element", $"Opaque presentation element {element.Id} requires its validated source package and cannot be authored from scratch.");
                shapeTree.Append(BuildShape(element, nativeId++, hyperlinkContext));
            }
            slidePart.Slide.Save();
        }
        presentationPart.Presentation = new P.Presentation(
            new P.SlideMasterIdList(new P.SlideMasterId { Id = 2_147_483_648U, RelationshipId = "rIdMaster1" }),
            slideIdList,
            new P.SlideSize
            {
                Cx = checked((int)(artifact.SlideWidthEmu > 0 ? artifact.SlideWidthEmu : DefaultSlideWidthEmu)),
                Cy = checked((int)(artifact.SlideHeightEmu > 0 ? artifact.SlideHeightEmu : DefaultSlideHeightEmu)),
            },
            new P.NotesSize { Cx = 6_858_000L, Cy = 9_144_000L },
            new P.DefaultTextStyle());
        themePart.Theme.Save();
        layoutPart.SlideLayout.Save();
        masterPart.SlideMaster.Save();
        presentationPart.Presentation.Save();
    }

    private static P.Shape BuildShape(PresentationElement source, uint nativeId, PptxHyperlinkContext hyperlinkContext)
    {
        var semantic = source.Shape;
        var properties = new P.ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = semantic.LeftEmu, Y = semantic.TopEmu },
                new A.Extents { Cx = semantic.WidthEmu, Cy = semantic.HeightEmu }),
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = semantic.Geometry == "ellipse" ? A.ShapeTypeValues.Ellipse : A.ShapeTypeValues.Rectangle });
        properties.Append(string.IsNullOrWhiteSpace(semantic.FillRgb)
            ? new A.NoFill()
            : new A.SolidFill(new A.RgbColorModelHex { Val = PptxColor.Normalize(semantic.FillRgb) }));
        var outline = new A.Outline { Width = checked((int)semantic.LineWidthEmu) };
        outline.Append(string.IsNullOrWhiteSpace(semantic.LineRgb)
            ? new A.NoFill()
            : new A.SolidFill(new A.RgbColorModelHex { Val = PptxColor.Normalize(semantic.LineRgb) }));
        properties.Append(outline);
        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = nativeId, Name = source.Name },
                new P.NonVisualShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            properties,
            PptxTextCodec.Build(semantic, hyperlinkContext));
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

    private static (long Left, long Top, long Width, long Height) ReadFrame(OpenXmlElement element)
    {
        var transform = element.Descendants<A.Transform2D>().FirstOrDefault();
        if (transform?.Offset is not null && transform.Extents is not null)
            return (transform.Offset.X?.Value ?? 0, transform.Offset.Y?.Value ?? 0, transform.Extents.Cx?.Value ?? 0, transform.Extents.Cy?.Value ?? 0);
        var offset = element.Descendants<A.Offset>().FirstOrDefault();
        var extents = element.Descendants<A.Extents>().FirstOrDefault();
        return (offset?.X?.Value ?? 0, offset?.Y?.Value ?? 0, extents?.Cx?.Value ?? 0, extents?.Cy?.Value ?? 0);
    }

    private static string Geometry(P.ShapeProperties? properties)
    {
        var value = properties?.GetFirstChild<A.PresetGeometry>()?.Preset?.Value;
        if (value is null) return "rect";
        return value.Equals(A.ShapeTypeValues.Ellipse) ? "ellipse" : value.Equals(A.ShapeTypeValues.Rectangle) ? "rect" : value.ToString() ?? "rect";
    }

    private static string ElementName(OpenXmlElement element, int index) =>
        element.Descendants<P.NonVisualDrawingProperties>().FirstOrDefault()?.Name?.Value ?? $"{element.LocalName} {index + 1}";

    private static string DescendantText(OpenXmlElement? element) =>
        element is null ? string.Empty : string.Concat(element.Descendants<A.Text>().Select(text => text.Text));

    private static string SemanticHash(PresentationElement element)
    {
        var semantic = element.Clone();
        semantic.Id = string.Empty;
        semantic.Source = null;
        if (semantic.ContentCase == PresentationElement.ContentOneofCase.Shape) PptxTextCodec.NormalizeSemantics(semantic.Shape);
        return Hash(semantic.ToByteArray());
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
    private static string HashElement(OpenXmlElement element) => Hash(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static void ValidateEnvelope(ArtifactEnvelope envelope, EffectiveCodecLimits limits)
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

        ulong items = 0;
        foreach (var slide in envelope.Presentation.Slides)
        {
            foreach (var element in slide.Elements)
            {
                items++;
                if (items > limits.MaxCells)
                    throw new CodecException("presentation_item_budget_exceeded", $"Presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).");
                if (element.ContentCase == PresentationElement.ContentOneofCase.Shape)
                {
                    if (element.Shape.LeftEmu < 0 || element.Shape.TopEmu < 0 || element.Shape.WidthEmu <= 0 || element.Shape.HeightEmu <= 0 || element.Shape.LineWidthEmu < 0 || element.Shape.LineWidthEmu > int.MaxValue)
                        throw new CodecException("invalid_presentation_frame", $"Presentation shape {element.Id} has an invalid frame.");
                    if (element.Shape.Geometry is not ("rect" or "ellipse"))
                        throw new CodecException("unsupported_presentation_geometry", $"Presentation shape {element.Id} uses unsupported geometry {element.Shape.Geometry}.");
                    if (!string.IsNullOrWhiteSpace(element.Shape.FillRgb)) PptxColor.Normalize(element.Shape.FillRgb);
                    if (!string.IsNullOrWhiteSpace(element.Shape.LineRgb)) PptxColor.Normalize(element.Shape.LineRgb);
                    PptxTextCodec.Validate(element.Shape);
                }
                else if (element.ContentCase != PresentationElement.ContentOneofCase.Opaque)
                    throw new CodecException("missing_presentation_element_content", $"Presentation element {element.Id} has no content.");
            }
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

    private static void ValidatePreservedSlideElements(byte[] sourceBytes, byte[] outputBytes, PresentationArtifact requested)
    {
        using var sourceStream = new MemoryStream(sourceBytes, writable: false);
        using var outputStream = new MemoryStream(outputBytes, writable: false);
        using var sourcePackage = PresentationDocument.Open(sourceStream, isEditable: false);
        using var outputPackage = PresentationDocument.Open(outputStream, isEditable: false);
        var sourceSlides = OrderedSlideParts(sourcePackage);
        var outputSlides = OrderedSlideParts(outputPackage);
        if (sourceSlides.Length != requested.Slides.Count || outputSlides.Length != requested.Slides.Count)
            throw new CodecException("presentation_postwrite_topology_changed", "PPTX slide topology changed during source-preserving export.");
        var sourceIdByPartPath = sourceSlides
            .Select((part, index) => (Path: PartPath(part), Id: requested.Slides[index].Id))
            .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
        var outputIdByPartPath = outputSlides
            .Select((part, index) => (Path: PartPath(part), Id: requested.Slides[index].Id))
            .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);

        for (var slideIndex = 0; slideIndex < requested.Slides.Count; slideIndex++)
        {
            var sourceHyperlinks = new PptxHyperlinkContext(sourceSlides[slideIndex], sourceIdByPartPath);
            var outputHyperlinks = new PptxHyperlinkContext(outputSlides[slideIndex], outputIdByPartPath);
            var before = ShapeElements(sourceSlides[slideIndex].Slide!.CommonSlideData!.ShapeTree!);
            var after = ShapeElements(outputSlides[slideIndex].Slide!.CommonSlideData!.ShapeTree!);
            var elements = requested.Slides[slideIndex].Elements;
            if (before.Length != elements.Count || after.Length != elements.Count)
                throw new CodecException("presentation_postwrite_topology_changed", $"PPTX slide {slideIndex + 1} element topology changed during source-preserving export.", PartPath(outputSlides[slideIndex]));
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
                if (before[elementIndex] is not P.Shape beforeShape || after[elementIndex] is not P.Shape afterShape)
                    throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} edited element {elementIndex + 1} changed native element type.", PartPath(outputSlides[slideIndex]));
                if (!ShapeResidualHash(beforeShape, sourceHyperlinks).Equals(ShapeResidualHash(afterShape, outputHyperlinks), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_unmodeled_shape_content_changed",
                        $"PPTX slide {slideIndex + 1} edited shape {elementIndex + 1} changed unmodeled native content.",
                        PartPath(outputSlides[slideIndex]));
                var outputSemantic = ReadElement(afterShape, slideIndex, elementIndex, outputHyperlinks);
                if (!SemanticHash(outputSemantic).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_postwrite_semantics_mismatch",
                        $"PPTX slide {slideIndex + 1} edited shape {elementIndex + 1} does not match requested semantics after export.",
                        PartPath(outputSlides[slideIndex]));
            }
        }
    }

    private static SlidePart[] OrderedSlideParts(PresentationDocument package)
    {
        var presentationPart = package.PresentationPart ?? throw new CodecException("missing_presentation_part", "PPTX package has no Presentation part.", "ppt/presentation.xml");
        return ResolveSlideParts(presentationPart, presentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>() ?? []);
    }

    private static SlidePart[] ResolveSlideParts(PresentationPart presentationPart, IEnumerable<P.SlideId> slideIds) =>
        slideIds.Select(slideId => presentationPart.GetPartById(slideId.RelationshipId?.Value ?? string.Empty) as SlidePart ??
            throw new CodecException("missing_slide_part", "PPTX presentation contains an unresolved slide relationship.", "ppt/presentation.xml"))
        .ToArray();

    private static string ShapeResidualHash(P.Shape source, PptxHyperlinkContext hyperlinkContext)
    {
        var shape = (P.Shape)source.CloneNode(true);
        if (shape.NonVisualShapeProperties?.NonVisualDrawingProperties is { } nonVisual) nonVisual.Name = string.Empty;
        if (shape.ShapeProperties is { } properties)
        {
            if (properties.Transform2D is { } transform)
            {
                if (transform.Offset is { } offset) { offset.X = 0L; offset.Y = 0L; }
                if (transform.Extents is { } extents) { extents.Cx = 1L; extents.Cy = 1L; }
            }
            if (properties.GetFirstChild<A.PresetGeometry>() is { } geometry) geometry.Preset = A.ShapeTypeValues.Rectangle;
            foreach (var fill in properties.ChildElements.Where(child => child is A.NoFill or A.SolidFill).ToArray()) fill.Remove();
            if (properties.GetFirstChild<A.Outline>() is { } outline)
            {
                outline.Width = 0;
                foreach (var fill in outline.ChildElements.Where(child => child is A.NoFill or A.SolidFill).ToArray()) fill.Remove();
            }
        }
        PptxTextCodec.ScrubModeledContent(shape.TextBody, hyperlinkContext);
        return HashElement(shape);
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

    private static void ValidateOffice2021(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var package = PresentationDocument.Open(stream, isEditable: false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package).Take(8).ToArray();
        if (errors.Length == 0) return;
        var detail = string.Join("; ", errors.Select(error => $"{error.Path?.XPath ?? error.Part?.Uri.ToString() ?? "package"}: {error.Description}"));
        throw new CodecException("openxml_validation_failed", $"Generated PPTX is not valid Office 2021 Open XML: {detail}");
    }
}
