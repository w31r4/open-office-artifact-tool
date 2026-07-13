using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Xunit;

namespace OpenOffice.OpenXmlCodec.Tests;

public sealed class PptxCodecTests
{
    [Fact]
    public void ProtocolRoundTripsMinimalPresentation()
    {
        var exported = Invoke(ExportRequest());
        Assert.True(exported.Ok, Diagnostics(exported));
        Assert.Equal("PK", System.Text.Encoding.ASCII.GetString(exported.File.Span[..2]));
        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var presentation = PresentationDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(presentation));

        var imported = Import(exported.File.ToByteArray());
        Assert.True(imported.Ok, Diagnostics(imported));
        Assert.Equal(ArtifactFamily.Presentation, imported.Artifact.Family);
        Assert.Collection(imported.Artifact.Presentation.Slides,
            slide => Assert.Collection(slide.Elements,
                shape =>
                {
                    Assert.Equal("Quarterly brief", shape.Shape.Text);
                    Assert.True(shape.Source.Editable);
                    Assert.Equal("rect", shape.Shape.Geometry);
                }));
    }

    [Fact]
    public void SourcePreservingExportEditsSimpleShapeAndKeepsPictureGraph()
    {
        var first = Invoke(ExportRequest());
        var source = AddPicture(first.File.ToByteArray());
        var imported = Import(source);
        Assert.True(imported.Ok, Diagnostics(imported));
        var slide = Assert.Single(imported.Artifact.Presentation.Slides);
        Assert.Equal(2, slide.Elements.Count);
        Assert.True(slide.Elements[0].Source.Editable);
        Assert.False(slide.Elements[1].Source.Editable);
        Assert.Equal("pic", slide.Elements[1].Opaque.ElementName);
        Assert.Contains(imported.Artifact.OpaqueOpc.Parts, part => part.Path.EndsWith("/image.png", StringComparison.OrdinalIgnoreCase) || part.Path.EndsWith("/image1.png", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(imported.Artifact.OpaqueOpc.PackageRelationships,
            relationship => relationship.SourcePath == "ppt/slides/slide1.xml" && relationship.Id == "rIdImage1");

        slide.Elements[0].Shape.Text = "Edited safely";
        var preserved = Export(imported.Artifact);
        Assert.True(preserved.Ok, Diagnostics(preserved));
        Assert.Equal("opaque_content_preserved", Assert.Single(preserved.Diagnostics).Code);
        using (var stream = new MemoryStream(preserved.File.ToByteArray()))
        using (var presentation = PresentationDocument.Open(stream, false))
        {
            var slidePart = presentation.PresentationPart!.SlideParts.Single();
            Assert.Equal("Edited safely", string.Concat(slidePart.Slide!.Descendants<A.Text>().Take(1).Select(text => text.Text)));
            Assert.Single(slidePart.ImageParts);
            Assert.NotNull(slidePart.Slide.Descendants<P.Picture>().SingleOrDefault());
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(presentation));
        }

        slide.Elements[1].Opaque.Text = "Unsafe picture edit";
        var rejected = Export(imported.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_presentation_edit", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void SourcePreservingExportRejectsTamperedBindingAndTopology()
    {
        var imported = Import(Invoke(ExportRequest()).File.ToByteArray());
        imported.Artifact.Presentation.Slides[0].Elements[0].Source.SemanticSha256 = new string('0', 64);
        var tampered = Export(imported.Artifact);
        Assert.False(tampered.Ok);
        Assert.Equal("presentation_source_semantics_mismatch", Assert.Single(tampered.Diagnostics).Code);

        imported = Import(Invoke(ExportRequest()).File.ToByteArray());
        imported.Artifact.Presentation.Slides[0].Elements.Clear();
        var changed = Export(imported.Artifact);
        Assert.False(changed.Ok);
        Assert.Equal("presentation_element_topology_changed", Assert.Single(changed.Diagnostics).Code);
    }

    [Fact]
    public void ProtocolReturnsStructuredSlideAndItemBudgetFailures()
    {
        var exported = Invoke(ExportRequest());
        var slideBudget = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportPptx,
            Family = ArtifactFamily.Presentation,
            File = exported.File,
            Limits = new CodecLimits { MaxSheets = 0, MaxCells = 1 },
        });
        Assert.True(slideBudget.Ok, Diagnostics(slideBudget));

        var itemBudget = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ImportPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(AddPicture(exported.File.ToByteArray())),
            Limits = new CodecLimits { MaxCells = 1 },
        });
        Assert.False(itemBudget.Ok);
        Assert.Equal("presentation_item_budget_exceeded", Assert.Single(itemBudget.Diagnostics).Code);
    }

    private static CodecResponse Invoke(CodecRequest request) =>
        CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));

    private static CodecResponse Import(byte[] bytes) => Invoke(new CodecRequest
    {
        ProtocolVersion = CodecProtocol.ProtocolVersion,
        Operation = CodecOperation.ImportPptx,
        Family = ArtifactFamily.Presentation,
        File = ByteString.CopyFrom(bytes),
    });

    private static CodecResponse Export(ArtifactEnvelope artifact) => Invoke(new CodecRequest
    {
        ProtocolVersion = CodecProtocol.ProtocolVersion,
        Operation = CodecOperation.ExportPptx,
        Family = ArtifactFamily.Presentation,
        Artifact = artifact,
    });

    private static string Diagnostics(CodecResponse response) =>
        string.Join("\n", response.Diagnostics.Select(item => $"{item.Code}: {item.Message}"));

    private static CodecRequest ExportRequest()
    {
        var slide = new PresentationSlide { Id = "presentation/slide/1", Name = "Overview" };
        slide.Elements.Add(new PresentationElement
        {
            Id = "presentation/slide/1/title",
            Name = "Title",
            Shape = new PresentationShape
            {
                Geometry = "rect",
                LeftEmu = 571_500,
                TopEmu = 381_000,
                WidthEmu = 8_191_500,
                HeightEmu = 666_750,
                Text = "Quarterly brief",
                FillRgb = "FFFFFF",
                LineRgb = "334155",
                LineWidthEmu = 9_525,
            },
        });
        var presentation = new PresentationArtifact
        {
            Id = "presentation/test",
            Name = "Quarterly brief",
            SlideWidthEmu = 12_192_000,
            SlideHeightEmu = 6_858_000,
        };
        presentation.Slides.Add(slide);
        return new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ExportPptx,
            Family = ArtifactFamily.Presentation,
            Artifact = new ArtifactEnvelope
            {
                ProtocolVersion = CodecProtocol.ProtocolVersion,
                Family = ArtifactFamily.Presentation,
                Presentation = presentation,
            },
        };
    }

    private static byte[] AddPicture(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var presentation = PresentationDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var slidePart = presentation.PresentationPart!.SlideParts.Single();
            var imagePart = slidePart.AddImagePart(ImagePartType.Png, "rIdImage1");
            using (var image = new MemoryStream(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=")))
                imagePart.FeedData(image);
            slidePart.Slide!.CommonSlideData!.ShapeTree!.Append(new P.Picture(
                new P.NonVisualPictureProperties(
                    new P.NonVisualDrawingProperties { Id = 3U, Name = "Preserved picture", Description = "Opaque package evidence" },
                    new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.BlipFill(
                    new A.Blip { Embed = "rIdImage1" },
                    new A.Stretch(new A.FillRectangle())),
                new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = 571_500L, Y = 1_524_000L },
                        new A.Extents { Cx = 1_714_500L, Cy = 1_143_000L }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })));
            slidePart.Slide.Save();
        }
        return stream.ToArray();
    }
}
