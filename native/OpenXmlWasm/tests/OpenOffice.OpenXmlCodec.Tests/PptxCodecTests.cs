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

    [Fact]
    public void RichTextRoundTripsAndPreservesUnmodeledRunAndParagraphProperties()
    {
        var authored = Invoke(RichTextExportRequest());
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = AddUnmodeledTextProperties(authored.File.ToByteArray());
        var imported = Import(source);
        Assert.True(imported.Ok, Diagnostics(imported));
        var shape = Assert.Single(Assert.Single(imported.Artifact.Presentation.Slides).Elements).Shape;
        Assert.Equal("Quarterly brief\nSource-bound detail\nExplicitly unbulleted\nOpaque picture marker", shape.Text);
        Assert.Equal(4, shape.TextBody.Paragraphs.Count);
        Assert.Equal("center", shape.TextBody.Paragraphs[0].Alignment);
        Assert.Equal(2, shape.TextBody.Paragraphs[0].Runs.Count);
        Assert.True(shape.TextBody.Paragraphs[0].Runs[0].Bold);
        Assert.Equal(27, shape.TextBody.Paragraphs[0].Runs[0].FontSizePoints);
        Assert.Equal("Aptos Display", shape.TextBody.Paragraphs[0].Runs[0].FontFamily);
        Assert.Equal("0F172A", shape.TextBody.Paragraphs[0].Runs[0].ColorRgb);
        Assert.True(shape.TextBody.Paragraphs[0].Runs[1].Italic);
        Assert.Equal(PresentationTextParagraph.BulletOneofCase.BulletCharacter, shape.TextBody.Paragraphs[0].BulletCase);
        Assert.Equal("•", shape.TextBody.Paragraphs[0].BulletCharacter);
        Assert.Equal("Georgia", shape.TextBody.Paragraphs[0].BulletFontFamily);
        Assert.Equal("DC2626", shape.TextBody.Paragraphs[0].BulletColorRgb);
        Assert.Equal(1.5, shape.TextBody.Paragraphs[0].BulletSizePercent);
        Assert.False(shape.TextBody.Paragraphs[1].HasAlignment);
        Assert.Equal("romanLcPeriod", shape.TextBody.Paragraphs[1].AutoNumber.Scheme);
        Assert.Equal(3U, shape.TextBody.Paragraphs[1].AutoNumber.StartAt);
        Assert.True(shape.TextBody.Paragraphs[1].BulletFontFollowText);
        Assert.True(shape.TextBody.Paragraphs[1].BulletColorFollowText);
        Assert.True(shape.TextBody.Paragraphs[1].BulletSizeFollowText);
        Assert.Equal(PresentationTextParagraph.BulletOneofCase.NoBullet, shape.TextBody.Paragraphs[2].BulletCase);
        Assert.Equal(18, shape.TextBody.Paragraphs[2].BulletSizePoints);
        Assert.Equal(PresentationTextParagraph.BulletOneofCase.None, shape.TextBody.Paragraphs[3].BulletCase);
        Assert.Equal("Wingdings", shape.TextBody.Paragraphs[3].BulletFontFamily);
        Assert.Equal(PresentationTextParagraph.BulletColorOneofCase.None, shape.TextBody.Paragraphs[3].BulletColorCase);

        shape.TextBody.Paragraphs[0].Runs[0].Text = "Updated ";
        shape.TextBody.Paragraphs[0].Runs[0].Bold = false;
        shape.TextBody.Paragraphs[0].Runs[0].ColorRgb = "2563EB";
        shape.TextBody.Paragraphs[0].BulletCharacter = "◆";
        shape.TextBody.Paragraphs[0].BulletFontFollowText = true;
        shape.TextBody.Paragraphs[0].BulletColorRgb = "2563EB";
        shape.TextBody.Paragraphs[0].BulletSizePoints = 18;
        shape.TextBody.Paragraphs[1].AutoNumber = new PresentationAutoNumberBullet { Scheme = "arabicPeriod", StartAt = 5 };
        shape.TextBody.Paragraphs[1].BulletFontFamily = "Aptos";
        shape.TextBody.Paragraphs[1].BulletColorRgb = "16A34A";
        shape.TextBody.Paragraphs[1].BulletSizePercent = 1.25;
        shape.TextBody.Paragraphs[2].BulletSizeFollowText = true;
        shape.TextBody.Paragraphs[3].BulletFontFamily = "Georgia";
        shape.Text = PptxTextCodec.Flatten(shape.TextBody);
        var preserved = Export(imported.Artifact);
        Assert.True(preserved.Ok, Diagnostics(preserved));
        using (var stream = new MemoryStream(preserved.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var nativeShape = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<P.Shape>().Single();
            var paragraph = nativeShape.TextBody!.Elements<A.Paragraph>().First();
            var run = paragraph.Elements<A.Run>().First();
            Assert.Equal("Updated ", run.Text!.Text);
            Assert.False(run.RunProperties!.Bold!.Value);
            Assert.Equal("2563EB", run.RunProperties.GetFirstChild<A.SolidFill>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal(A.TextUnderlineValues.Single, run.RunProperties.Underline!.Value);
            Assert.Equal("Noto Sans CJK SC", run.RunProperties.GetFirstChild<A.EastAsianFont>()!.Typeface!.Value);
            Assert.Equal("◆", paragraph.ParagraphProperties!.GetFirstChild<A.CharacterBullet>()!.Char!.Value);
            Assert.NotNull(paragraph.ParagraphProperties.GetFirstChild<A.BulletFontText>());
            Assert.Equal("2563EB", paragraph.ParagraphProperties.GetFirstChild<A.BulletColor>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal(1800, paragraph.ParagraphProperties.GetFirstChild<A.BulletSizePoints>()!.Val!.Value);
            var secondParagraph = nativeShape.TextBody.Elements<A.Paragraph>().ElementAt(1).ParagraphProperties!;
            var autoNumber = secondParagraph.GetFirstChild<A.AutoNumberedBullet>()!;
            Assert.Equal("arabicPeriod", autoNumber.Type!.InnerText);
            Assert.Equal(5, autoNumber.StartAt!.Value);
            Assert.Equal("Aptos", secondParagraph.GetFirstChild<A.BulletFont>()!.Typeface!.Value);
            Assert.Equal("16A34A", secondParagraph.GetFirstChild<A.BulletColor>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal(125000, secondParagraph.GetFirstChild<A.BulletSizePercentage>()!.Val!.Value);
            var thirdParagraph = nativeShape.TextBody.Elements<A.Paragraph>().ElementAt(2).ParagraphProperties!;
            Assert.NotNull(thirdParagraph.GetFirstChild<A.NoBullet>());
            Assert.NotNull(thirdParagraph.GetFirstChild<A.BulletSizeText>());
            var pictureParagraph = nativeShape.TextBody.Elements<A.Paragraph>().ElementAt(3).ParagraphProperties!;
            Assert.NotNull(pictureParagraph.GetFirstChild<A.PictureBullet>());
            Assert.Equal("Georgia", pictureParagraph.GetFirstChild<A.BulletFont>()!.Typeface!.Value);
            Assert.Equal(A.SchemeColorValues.Accent1, pictureParagraph.GetFirstChild<A.BulletColor>()!.GetFirstChild<A.SchemeColor>()!.Val!.Value);
            Assert.Single(package.PresentationPart.SlideParts.Single().ImageParts);
            Assert.Equal(A.TextAlignmentTypeValues.Distributed, secondParagraph.Alignment!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var pictureEdit = Import(source);
        var pictureShape = Assert.Single(Assert.Single(pictureEdit.Artifact.Presentation.Slides).Elements).Shape;
        pictureShape.TextBody.Paragraphs[3].BulletCharacter = "◆";
        pictureShape.Text = PptxTextCodec.Flatten(pictureShape.TextBody);
        var pictureRejected = Export(pictureEdit.Artifact);
        Assert.False(pictureRejected.Ok);
        Assert.Equal("unsupported_presentation_edit", Assert.Single(pictureRejected.Diagnostics).Code);

        var unknownColorEdit = Import(source);
        var unknownColorShape = Assert.Single(Assert.Single(unknownColorEdit.Artifact.Presentation.Slides).Elements).Shape;
        unknownColorShape.TextBody.Paragraphs[3].BulletColorRgb = "FF0000";
        unknownColorShape.Text = PptxTextCodec.Flatten(unknownColorShape.TextBody);
        var unknownColorRejected = Export(unknownColorEdit.Artifact);
        Assert.False(unknownColorRejected.Ok);
        Assert.Equal("unsupported_presentation_edit", Assert.Single(unknownColorRejected.Diagnostics).Code);

        shape.TextBody.Paragraphs[0].Runs.Add(new PresentationTextRun { Text = "unsafe topology" });
        shape.Text = PptxTextCodec.Flatten(shape.TextBody);
        var rejected = Export(imported.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("presentation_text_topology_changed", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void RunHyperlinksRoundTripAndEditOwnedSlideRelationships()
    {
        var authored = Invoke(HyperlinkExportRequest());
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var slidePart = OrderedSlides(package)[0];
            var links = slidePart.Slide!.Descendants<A.HyperlinkOnClick>().ToArray();
            Assert.Equal(4, links.Length);
            Assert.Equal("https://example.com/guide?x=1&y=2", Assert.Single(slidePart.HyperlinkRelationships).Uri.OriginalString);
            Assert.Equal("ppaction://hlinksldjump", links[1].Action!.Value);
            Assert.IsType<SlidePart>(slidePart.GetPartById(links[1].Id!.Value!));
            Assert.Equal("ppaction://hlinkshowjump?jump=nextslide", links[2].Action!.Value);
            Assert.Equal(string.Empty, links[2].Id!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var imported = Import(authored.File.ToByteArray());
        Assert.True(imported.Ok, Diagnostics(imported));
        var runs = imported.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].Runs;
        Assert.Equal("https://example.com/guide?x=1&y=2", runs[0].RunHyperlink.Uri);
        Assert.Equal("Read the guide", runs[0].RunHyperlink.Tooltip);
        Assert.True(runs[0].RunHyperlink.HasHistory);
        Assert.False(runs[0].RunHyperlink.History);
        Assert.True(runs[0].RunHyperlink.HighlightClick);
        Assert.Equal("presentation/slide/2", runs[1].RunHyperlink.SlideId);
        Assert.Equal("nextSlide", runs[2].RunHyperlink.Action);

        runs[0].RunHyperlink = new PresentationRunHyperlink { Uri = "https://example.com/updated", TargetFrame = "_blank" };
        runs[1].RunHyperlink = new PresentationRunHyperlink { SlideId = "presentation/slide/3", Tooltip = "Third slide" };
        runs[2].RunHyperlink = new PresentationRunHyperlink { Action = "lastSlide", HighlightClick = false };
        runs[3].NoHyperlink = true;
        var preserved = Export(imported.Artifact);
        Assert.True(preserved.Ok, Diagnostics(preserved));
        using (var stream = new MemoryStream(preserved.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var slides = OrderedSlides(package);
            var links = slides[0].Slide!.Descendants<A.HyperlinkOnClick>().ToArray();
            Assert.Equal(3, links.Length);
            Assert.Equal("https://example.com/updated", slides[0].HyperlinkRelationships.Single(link => link.Id == links[0].Id).Uri.OriginalString);
            Assert.Same(slides[2], slides[0].GetPartById(links[1].Id!.Value!));
            Assert.Equal("ppaction://hlinkshowjump?jump=lastslide", links[2].Action!.Value);
            Assert.False(links[2].HighlightClick!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }
        var reimported = Import(preserved.File.ToByteArray());
        var reimportedRuns = reimported.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].Runs;
        Assert.Equal("https://example.com/updated", reimportedRuns[0].RunHyperlink.Uri);
        Assert.Equal("presentation/slide/3", reimportedRuns[1].RunHyperlink.SlideId);
        Assert.Equal("lastSlide", reimportedRuns[2].RunHyperlink.Action);
        Assert.Equal(PresentationTextRun.HyperlinkOneofCase.None, reimportedRuns[3].HyperlinkCase);
    }

    [Fact]
    public void UnknownRunClickActionsArePreservedAndCannotBeReplaced()
    {
        var authored = Invoke(HyperlinkExportRequest());
        var source = AddUnknownRunClick(authored.File.ToByteArray());
        var imported = Import(source);
        Assert.True(imported.Ok, Diagnostics(imported));
        var run = imported.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].Runs[0];
        Assert.Equal(PresentationTextRun.HyperlinkOneofCase.None, run.HyperlinkCase);
        run.Text = "Edited text";
        imported.Artifact.Presentation.Slides[0].Elements[0].Shape.Text = PptxTextCodec.Flatten(imported.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody);
        var preserved = Export(imported.Artifact);
        Assert.True(preserved.Ok, Diagnostics(preserved));
        using (var stream = new MemoryStream(preserved.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
            Assert.Equal("ppaction://customshow?id=99", OrderedSlides(package)[0].Slide!.Descendants<A.HyperlinkOnClick>().First().Action!.Value);

        imported = Import(source);
        run = imported.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].Runs[0];
        run.RunHyperlink = new PresentationRunHyperlink { Uri = "https://example.com/replacement" };
        var rejected = Export(imported.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_presentation_edit", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void InvalidRunHyperlinksFailClosed()
    {
        var request = HyperlinkExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].Runs[0].NoHyperlink = false;
        var invalidNone = Invoke(request);
        Assert.False(invalidNone.Ok);
        Assert.Equal("invalid_presentation_hyperlink", Assert.Single(invalidNone.Diagnostics).Code);

        request = HyperlinkExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].Runs[0].RunHyperlink = new PresentationRunHyperlink { Uri = "javascript:alert(1)" };
        var forbiddenUri = Invoke(request);
        Assert.False(forbiddenUri.Ok);
        Assert.Equal("invalid_presentation_hyperlink", Assert.Single(forbiddenUri.Diagnostics).Code);

        request = HyperlinkExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].Runs[0].RunHyperlink = new PresentationRunHyperlink { Action = "customShow" };
        var invalidAction = Invoke(request);
        Assert.False(invalidAction.Ok);
        Assert.Equal("invalid_presentation_hyperlink", Assert.Single(invalidAction.Diagnostics).Code);

        request = HyperlinkExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].Runs[0].RunHyperlink = new PresentationRunHyperlink { SlideId = "presentation/slide/missing" };
        var missingSlide = Invoke(request);
        Assert.False(missingSlide.Ok);
        Assert.Equal("invalid_presentation_hyperlink", Assert.Single(missingSlide.Diagnostics).Code);
    }

    [Fact]
    public void InvalidListMarkersAndStylesFailClosed()
    {
        var request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].BulletCharacter = "two";
        var invalidCharacter = Invoke(request);
        Assert.False(invalidCharacter.Ok);
        Assert.Equal("invalid_presentation_text", Assert.Single(invalidCharacter.Diagnostics).Code);

        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[1].AutoNumber = new PresentationAutoNumberBullet { Scheme = "not-a-scheme" };
        var invalidScheme = Invoke(request);
        Assert.False(invalidScheme.Ok);
        Assert.Equal("invalid_presentation_text", Assert.Single(invalidScheme.Diagnostics).Code);

        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[2].NoBullet = false;
        var invalidNone = Invoke(request);
        Assert.False(invalidNone.Ok);
        Assert.Equal("invalid_presentation_text", Assert.Single(invalidNone.Diagnostics).Code);

        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].BulletFontFamily = " ";
        var invalidFont = Invoke(request);
        Assert.False(invalidFont.Ok);
        Assert.Equal("invalid_presentation_text", Assert.Single(invalidFont.Diagnostics).Code);

        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].BulletColorRgb = "not-rgb";
        var invalidColor = Invoke(request);
        Assert.False(invalidColor.Ok);
        Assert.Equal("invalid_presentation_color", Assert.Single(invalidColor.Diagnostics).Code);

        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].BulletSizePoints = 0.5;
        var invalidSize = Invoke(request);
        Assert.False(invalidSize.Ok);
        Assert.Equal("invalid_presentation_text", Assert.Single(invalidSize.Diagnostics).Code);

        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[1].BulletSizeFollowText = false;
        var invalidFollow = Invoke(request);
        Assert.False(invalidFollow.Ok);
        Assert.Equal("invalid_presentation_text", Assert.Single(invalidFollow.Diagnostics).Code);
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

    private static CodecRequest RichTextExportRequest()
    {
        var first = new PresentationTextParagraph
        {
            Alignment = "center",
            BulletCharacter = "•",
            BulletFontFamily = "Georgia",
            BulletColorRgb = "DC2626",
            BulletSizePercent = 1.5,
        };
        first.Runs.Add(new PresentationTextRun
        {
            Text = "Quarterly ",
            Bold = true,
            FontSizePoints = 27,
            FontFamily = "Aptos Display",
            ColorRgb = "0F172A",
        });
        first.Runs.Add(new PresentationTextRun { Text = "brief", Italic = true, FontSizePoints = 27 });
        var second = new PresentationTextParagraph
        {
            Level = 1,
            AutoNumber = new PresentationAutoNumberBullet { Scheme = "romanLcPeriod", StartAt = 3 },
            BulletFontFollowText = true,
            BulletColorFollowText = true,
            BulletSizeFollowText = true,
        };
        second.Runs.Add(new PresentationTextRun { Text = "Source-bound detail", FontSizePoints = 15 });
        var third = new PresentationTextParagraph { NoBullet = true, BulletSizePoints = 18 };
        third.Runs.Add(new PresentationTextRun { Text = "Explicitly unbulleted", FontSizePoints = 15 });
        var fourth = new PresentationTextParagraph();
        fourth.Runs.Add(new PresentationTextRun { Text = "Opaque picture marker", FontSizePoints = 15 });
        var textBody = new PresentationTextBody();
        textBody.Paragraphs.Add(first);
        textBody.Paragraphs.Add(second);
        textBody.Paragraphs.Add(third);
        textBody.Paragraphs.Add(fourth);
        var slide = new PresentationSlide { Id = "presentation/slide/1", Name = "Rich text" };
        slide.Elements.Add(new PresentationElement
        {
            Id = "presentation/slide/1/rich-text",
            Name = "Rich text",
            Shape = new PresentationShape
            {
                Geometry = "rect",
                LeftEmu = 571_500,
                TopEmu = 381_000,
                WidthEmu = 8_191_500,
                HeightEmu = 1_714_500,
                Text = PptxTextCodec.Flatten(textBody),
                TextBody = textBody,
                FillRgb = "FFFFFF",
                LineRgb = "334155",
                LineWidthEmu = 9_525,
            },
        });
        var presentation = new PresentationArtifact
        {
            Id = "presentation/rich-text",
            Name = "Rich text",
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

    private static CodecRequest HyperlinkExportRequest()
    {
        var paragraph = new PresentationTextParagraph();
        paragraph.Runs.Add(new PresentationTextRun
        {
            Text = "Guide ",
            RunHyperlink = new PresentationRunHyperlink
            {
                Uri = "https://example.com/guide?x=1&y=2",
                Tooltip = "Read the guide",
                TargetFrame = "_blank",
                History = false,
                HighlightClick = true,
            },
        });
        paragraph.Runs.Add(new PresentationTextRun
        {
            Text = "Details ",
            RunHyperlink = new PresentationRunHyperlink { SlideId = "presentation/slide/2" },
        });
        paragraph.Runs.Add(new PresentationTextRun
        {
            Text = "Next ",
            RunHyperlink = new PresentationRunHyperlink { Action = "nextSlide" },
        });
        paragraph.Runs.Add(new PresentationTextRun
        {
            Text = "End",
            RunHyperlink = new PresentationRunHyperlink { Action = "endShow" },
        });
        var body = new PresentationTextBody();
        body.Paragraphs.Add(paragraph);
        var first = new PresentationSlide { Id = "presentation/slide/1", Name = "Links" };
        first.Elements.Add(new PresentationElement
        {
            Id = "presentation/slide/1/links",
            Name = "Links",
            Shape = new PresentationShape
            {
                Geometry = "rect",
                LeftEmu = 571_500,
                TopEmu = 381_000,
                WidthEmu = 8_191_500,
                HeightEmu = 666_750,
                Text = PptxTextCodec.Flatten(body),
                TextBody = body,
                FillRgb = "FFFFFF",
                LineRgb = "334155",
                LineWidthEmu = 9_525,
            },
        });
        var presentation = new PresentationArtifact
        {
            Id = "presentation/hyperlinks",
            Name = "Hyperlinks",
            SlideWidthEmu = 12_192_000,
            SlideHeightEmu = 6_858_000,
        };
        presentation.Slides.Add(first);
        presentation.Slides.Add(new PresentationSlide { Id = "presentation/slide/2", Name = "Details" });
        presentation.Slides.Add(new PresentationSlide { Id = "presentation/slide/3", Name = "Appendix" });
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

    private static SlidePart[] OrderedSlides(PresentationDocument package)
    {
        var presentationPart = package.PresentationPart!;
        return presentationPart.Presentation!.SlideIdList!.Elements<P.SlideId>()
            .Select(slideId => (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!.Value!))
            .ToArray();
    }

    private static byte[] AddUnknownRunClick(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var package = PresentationDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var click = OrderedSlides(package)[0].Slide!.Descendants<A.HyperlinkOnClick>().First();
            click.Id = string.Empty;
            click.Action = "ppaction://customshow?id=99";
        }
        return stream.ToArray();
    }

    private static byte[] AddUnmodeledTextProperties(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var presentation = PresentationDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var shape = presentation.PresentationPart!.SlideParts.Single().Slide!.Descendants<P.Shape>().Single();
            var slidePart = presentation.PresentationPart!.SlideParts.Single();
            var paragraph = shape.TextBody!.Elements<A.Paragraph>().First();
            shape.TextBody.Elements<A.Paragraph>().ElementAt(1).ParagraphProperties!.Alignment = A.TextAlignmentTypeValues.Distributed;
            var pictureParagraph = shape.TextBody.Elements<A.Paragraph>().Last();
            var pictureProperties = pictureParagraph.ParagraphProperties ?? pictureParagraph.PrependChild(new A.ParagraphProperties());
            pictureProperties.AddChild(new A.BulletColor(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }), true);
            pictureProperties.AddChild(new A.BulletFont { Typeface = "Wingdings" }, true);
            pictureProperties.AddChild(new A.PictureBullet(new A.Blip { Embed = "rIdTextBullet1" }), true);
            var imagePart = slidePart.AddImagePart(ImagePartType.Png, "rIdTextBullet1");
            using (var image = new MemoryStream(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=")))
                imagePart.FeedData(image);
            var properties = paragraph.Elements<A.Run>().First().RunProperties!;
            properties.Underline = A.TextUnderlineValues.Single;
            properties.Append(new A.EastAsianFont { Typeface = "Noto Sans CJK SC" });
        }
        return stream.ToArray();
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
