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
    public void MasterGraphAndTextStylesAuthorImportEditDeleteAndPreserveResidualContent()
    {
        var request = ExportRequest();
        var styles = new PresentationMasterTextStyles();
        styles.TitleLevels.Add(new PresentationTextParagraph
        {
            Level = 0,
            Alignment = "center",
            DefaultRunProperties = new PresentationTextStyle
            {
                Bold = true,
                FontSizePoints = 30,
                FontFamily = "Aptos Display",
                ColorScheme = "accent1",
            },
        });
        styles.BodyLevels.Add(new PresentationTextParagraph
        {
            Level = 1,
            MarginLeftEmu = 685_800,
            IndentEmu = -228_600,
            BulletCharacter = "•",
        });
        request.Artifact.Presentation.Masters.Add(new PresentationMaster
        {
            Id = "master/authored",
            Name = "Authored Master",
            TextStyles = styles,
        });

        var authored = Invoke(request);
        Assert.True(authored.Ok, Diagnostics(authored));
        byte[] source;
        using (var stream = new MemoryStream())
        {
            stream.Write(authored.File.Span);
            stream.Position = 0;
            using (var package = PresentationDocument.Open(stream, true))
            {
                var master = package.PresentationPart!.SlideMasterParts.Single().SlideMaster!;
                master.TextStyles!.TitleStyle!.GetFirstChild<A.Level1ParagraphProperties>()!.RightMargin = 123_456;
                master.Save();
            }
            source = stream.ToArray();
        }

        var imported = Import(source);
        Assert.True(imported.Ok, Diagnostics(imported));
        var importedMaster = Assert.Single(imported.Artifact.Presentation.Masters);
        var importedLayout = Assert.Single(imported.Artifact.Presentation.Layouts);
        var importedSlide = Assert.Single(imported.Artifact.Presentation.Slides);
        Assert.Equal("presentation/master/1", importedMaster.Id);
        Assert.Equal("Authored Master", importedMaster.Name);
        Assert.True(importedMaster.Source.TextStylesEditable);
        Assert.Equal("presentation/master/1/layout/1", importedLayout.Id);
        Assert.Equal(importedMaster.Id, importedLayout.MasterId);
        Assert.Equal("blank", importedLayout.Type);
        Assert.Equal(importedLayout.Id, importedSlide.LayoutId);
        Assert.False(string.IsNullOrWhiteSpace(importedSlide.Source.LayoutRelationshipId));
        Assert.Equal("center", Assert.Single(importedMaster.TextStyles.TitleLevels).Alignment);
        Assert.Equal(1U, Assert.Single(importedMaster.TextStyles.BodyLevels).Level);

        importedMaster.TextStyles.TitleLevels[0].Alignment = "right";
        importedMaster.TextStyles.BodyLevels.Clear();
        importedMaster.TextStyles.DeletedBodyLevels.Add(1);
        importedMaster.TextStyles.OtherLevels.Add(new PresentationTextParagraph
        {
            Level = 2,
            PictureBullet = new PresentationPictureBullet { Uri = "https://example.com/master-marker.png" },
            BulletColorScheme = "accent3",
        });
        var preserved = Export(imported.Artifact);
        Assert.True(preserved.Ok, Diagnostics(preserved));
        using (var stream = new MemoryStream(preserved.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var masterPart = package.PresentationPart!.SlideMasterParts.Single();
            var master = masterPart.SlideMaster!;
            var title = master.TextStyles!.TitleStyle!.GetFirstChild<A.Level1ParagraphProperties>()!;
            Assert.Equal(A.TextAlignmentTypeValues.Right, title.Alignment!.Value);
            Assert.Equal(123_456, title.RightMargin!.Value);
            Assert.Null(master.TextStyles.BodyStyle!.GetFirstChild<A.Level2ParagraphProperties>());
            var other = master.TextStyles.OtherStyle!.GetFirstChild<A.Level3ParagraphProperties>()!;
            Assert.NotNull(other.GetFirstChild<A.PictureBullet>());
            Assert.Equal(A.SchemeColorValues.Accent3, other.GetFirstChild<A.BulletColor>()!.GetFirstChild<A.SchemeColor>()!.Val!.Value);
            Assert.Contains(masterPart.ExternalRelationships, relationship =>
                relationship.RelationshipType.EndsWith("/image", StringComparison.Ordinal) &&
                relationship.Uri.OriginalString == "https://example.com/master-marker.png");
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var roundTrip = Import(preserved.File.ToByteArray());
        Assert.True(roundTrip.Ok, Diagnostics(roundTrip));
        var roundTripStyles = Assert.Single(roundTrip.Artifact.Presentation.Masters).TextStyles;
        Assert.Equal("right", Assert.Single(roundTripStyles.TitleLevels).Alignment);
        Assert.Empty(roundTripStyles.BodyLevels);
        Assert.Equal(PresentationTextParagraph.BulletOneofCase.PictureBullet, Assert.Single(roundTripStyles.OtherLevels).BulletCase);

        imported.Artifact.Presentation.Slides[0].LayoutId = "presentation/master/1/layout/missing";
        var rebound = Export(imported.Artifact);
        Assert.False(rebound.Ok);
        Assert.Equal("invalid_presentation_layout", Assert.Single(rebound.Diagnostics).Code);

        request = ExportRequest();
        var invalidStyles = new PresentationMasterTextStyles();
        invalidStyles.TitleLevels.Add(new PresentationTextParagraph { Level = 0 });
        request.Artifact.Presentation.Masters.Add(new PresentationMaster { Id = "master/invalid", TextStyles = invalidStyles });
        var invalid = Invoke(request);
        Assert.False(invalid.Ok);
        Assert.Equal("invalid_presentation_master_style", Assert.Single(invalid.Diagnostics).Code);
    }

    [Fact]
    public void MasterAndLayoutBackgroundsAuthorImportEditAndRejectUnsupportedFill()
    {
        var request = ExportRequest();
        request.Artifact.Presentation.Masters.Add(new PresentationMaster
        {
            Id = "master/background",
            Name = "Background Master",
            Background = new PresentationBackground { ColorScheme = "accent1", StyleReferenceIndex = 1001 },
        });
        var authored = Invoke(request);
        Assert.True(authored.Ok, Diagnostics(authored));
        byte[] source;
        using (var stream = new MemoryStream())
        {
            stream.Write(authored.File.Span);
            stream.Position = 0;
            using (var package = PresentationDocument.Open(stream, true))
            {
                var master = package.PresentationPart!.SlideMasterParts.Single().SlideMaster!;
                var masterReference = master.CommonSlideData!.Background!.GetFirstChild<P.BackgroundStyleReference>()!;
                Assert.Equal(1001U, masterReference.Index!.Value);
                Assert.Equal(A.SchemeColorValues.Accent1, masterReference.GetFirstChild<A.SchemeColor>()!.Val!.Value);
                var layout = package.PresentationPart.SlideMasterParts.Single().SlideLayoutParts.Single().SlideLayout!;
                layout.CommonSlideData!.AddChild(new P.Background(
                    new P.BackgroundProperties(
                        new A.SolidFill(new A.RgbColorModelHex { Val = "FFF7ED" }),
                        new A.EffectList())), true);
                layout.Save();
            }
            source = stream.ToArray();
        }

        var imported = Import(source);
        Assert.True(imported.Ok, Diagnostics(imported));
        var masterArtifact = Assert.Single(imported.Artifact.Presentation.Masters);
        var layoutArtifact = Assert.Single(imported.Artifact.Presentation.Layouts);
        Assert.True(masterArtifact.Source.BackgroundEditable);
        Assert.Equal(PresentationBackground.KindOneofCase.StyleReferenceIndex, masterArtifact.Background.KindCase);
        Assert.True(layoutArtifact.Source.BackgroundEditable);
        Assert.Equal("FFF7ED", layoutArtifact.Background.ColorRgb);

        masterArtifact.Background = new PresentationBackground { ColorRgb = "112233", Solid = true };
        layoutArtifact.Background = new PresentationBackground { ColorScheme = "accent2", StyleReferenceIndex = 1002 };
        var edited = Export(imported.Artifact);
        Assert.True(edited.Ok, Diagnostics(edited));
        using (var stream = new MemoryStream(edited.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var master = package.PresentationPart!.SlideMasterParts.Single().SlideMaster!;
            Assert.Equal("112233", master.CommonSlideData!.Background!.Descendants<A.RgbColorModelHex>().Single().Val!.Value);
            var layout = package.PresentationPart.SlideMasterParts.Single().SlideLayoutParts.Single().SlideLayout!;
            var reference = layout.CommonSlideData!.Background!.GetFirstChild<P.BackgroundStyleReference>()!;
            Assert.Equal(1002U, reference.Index!.Value);
            Assert.Equal(A.SchemeColorValues.Accent2, reference.GetFirstChild<A.SchemeColor>()!.Val!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }
        var roundTrip = Import(edited.File.ToByteArray());
        Assert.Equal("112233", Assert.Single(roundTrip.Artifact.Presentation.Masters).Background.ColorRgb);
        Assert.Equal("accent2", Assert.Single(roundTrip.Artifact.Presentation.Layouts).Background.ColorScheme);

        byte[] unsupportedSource;
        using (var stream = new MemoryStream())
        {
            stream.Write(authored.File.Span);
            stream.Position = 0;
            using (var package = PresentationDocument.Open(stream, true))
            {
                var layout = package.PresentationPart!.SlideMasterParts.Single().SlideLayoutParts.Single().SlideLayout!;
                layout.CommonSlideData!.AddChild(new P.Background(
                    new P.BackgroundProperties(new A.NoFill(), new A.EffectList())), true);
                layout.Save();
            }
            unsupportedSource = stream.ToArray();
        }
        var unsupported = Import(unsupportedSource);
        Assert.True(unsupported.Ok, Diagnostics(unsupported));
        var unsupportedLayout = Assert.Single(unsupported.Artifact.Presentation.Layouts);
        Assert.False(unsupportedLayout.Source.BackgroundEditable);
        Assert.Null(unsupportedLayout.Background);
        unsupportedLayout.Background = new PresentationBackground { ColorRgb = "FFFFFF", Solid = true };
        var rejected = Export(unsupported.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_presentation_edit", Assert.Single(rejected.Diagnostics).Code);

        request = ExportRequest();
        request.Artifact.Presentation.Masters.Add(new PresentationMaster
        {
            Id = "master/invalid-background",
            Background = new PresentationBackground { ColorRgb = "FFFFFF", Solid = false },
        });
        var invalid = Invoke(request);
        Assert.False(invalid.Ok);
        Assert.Equal("invalid_presentation_background", Assert.Single(invalid.Diagnostics).Code);
    }

    [Fact]
    public void MasterAndLayoutPlaceholderTextRoundTripsWithOwnerLocalBindings()
    {
        var authored = Invoke(ExportRequest());
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = AddTemplatePlaceholders(authored.File.ToByteArray());
        var imported = Import(source);
        Assert.True(imported.Ok, Diagnostics(imported));
        var master = Assert.Single(imported.Artifact.Presentation.Masters);
        var layout = Assert.Single(imported.Artifact.Presentation.Layouts);
        var masterPlaceholder = Assert.Single(master.Placeholders);
        var layoutPlaceholder = Assert.Single(layout.Placeholders);
        Assert.Equal("title", masterPlaceholder.Type);
        Assert.Equal(0U, masterPlaceholder.Index);
        Assert.Equal("Master prompt", PptxTextCodec.Flatten(masterPlaceholder.TextBody));
        Assert.True(masterPlaceholder.Source.Editable);
        Assert.Equal("body", layoutPlaceholder.Type);
        Assert.Equal(2U, layoutPlaceholder.Index);
        Assert.Equal("Layout prompt", PptxTextCodec.Flatten(layoutPlaceholder.TextBody));
        Assert.True(layoutPlaceholder.Source.Editable);
        Assert.NotEqual(uint.MaxValue, masterPlaceholder.Source.ShapeTreeIndex);

        masterPlaceholder.TextBody.Paragraphs[0].Runs[0].Text = "Edited master prompt";
        layoutPlaceholder.TextBody.Paragraphs[0].Runs[0].Text = "Edited layout prompt";
        layoutPlaceholder.TextBody.Paragraphs[0].Runs[0].RunHyperlink = new PresentationRunHyperlink
        {
            Uri = "https://example.com/layout-help",
            Tooltip = "Layout help",
        };
        var edited = Export(imported.Artifact);
        Assert.True(edited.Ok, Diagnostics(edited));
        using (var stream = new MemoryStream(edited.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var masterPart = package.PresentationPart!.SlideMasterParts.Single();
            var nativeMaster = masterPart.SlideMaster!.CommonSlideData!.ShapeTree!.Elements<P.Shape>().Single();
            Assert.Equal("Edited master prompt", nativeMaster.Descendants<A.Text>().Single().Text);
            Assert.True(nativeMaster.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.GetFirstChild<P.PlaceholderShape>()!.HasCustomPrompt!.Value);
            Assert.Equal(60_000, nativeMaster.ShapeProperties!.Transform2D!.Rotation!.Value);
            var layoutPart = masterPart.SlideLayoutParts.Single();
            var nativeLayout = layoutPart.SlideLayout!.CommonSlideData!.ShapeTree!.Elements<P.Shape>().Single();
            Assert.Equal("Edited layout prompt", nativeLayout.Descendants<A.Text>().Single().Text);
            Assert.Contains(layoutPart.HyperlinkRelationships, relationship =>
                relationship.Uri.OriginalString == "https://example.com/layout-help");
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }
        var roundTrip = Import(edited.File.ToByteArray());
        Assert.Equal("Edited master prompt", PptxTextCodec.Flatten(Assert.Single(Assert.Single(roundTrip.Artifact.Presentation.Masters).Placeholders).TextBody));
        var roundTripLayoutRun = Assert.Single(Assert.Single(Assert.Single(roundTrip.Artifact.Presentation.Layouts).Placeholders).TextBody.Paragraphs).Runs.Single();
        Assert.Equal("Edited layout prompt", roundTripLayoutRun.Text);
        Assert.Equal("https://example.com/layout-help", roundTripLayoutRun.RunHyperlink.Uri);

        var topology = Import(source);
        Assert.Single(topology.Artifact.Presentation.Layouts).Placeholders.Clear();
        var topologyRejected = Export(topology.Artifact);
        Assert.False(topologyRejected.Ok);
        Assert.Equal("presentation_placeholder_topology_changed", Assert.Single(topologyRejected.Diagnostics).Code);

        var unsupported = Import(AddUnsupportedPlaceholderRun(source));
        var unsupportedPlaceholder = Assert.Single(Assert.Single(unsupported.Artifact.Presentation.Layouts).Placeholders);
        Assert.False(unsupportedPlaceholder.Source.Editable);
        unsupportedPlaceholder.TextBody.Paragraphs[0].Runs[0].Text = "Unsafe edit";
        var unsupportedRejected = Export(unsupported.Artifact);
        Assert.False(unsupportedRejected.Ok);
        Assert.Equal("unsupported_presentation_edit", Assert.Single(unsupportedRejected.Diagnostics).Code);
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
    public void SchemeBulletColorsAuthorImportAndEditWithoutFlatteningThemeIdentity()
    {
        var request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].BulletColorScheme = "accent1";
        var authored = Invoke(request);
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var color = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.BulletColor>().First();
            Assert.Equal(A.SchemeColorValues.Accent1, color.GetFirstChild<A.SchemeColor>()!.Val!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var imported = Import(authored.File.ToByteArray());
        Assert.True(imported.Ok, Diagnostics(imported));
        var paragraph = imported.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0];
        Assert.Equal(PresentationTextParagraph.BulletColorOneofCase.BulletColorScheme, paragraph.BulletColorCase);
        Assert.Equal("accent1", paragraph.BulletColorScheme);
        paragraph.BulletColorScheme = "tx2";
        var edited = Export(imported.Artifact);
        Assert.True(edited.Ok, Diagnostics(edited));
        using (var stream = new MemoryStream(edited.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var color = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.BulletColor>().First();
            Assert.Equal(A.SchemeColorValues.Text2, color.GetFirstChild<A.SchemeColor>()!.Val!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }
    }

    [Fact]
    public void ParagraphMarginsAndHangingIndentsAuthorImportEditAndDelete()
    {
        var request = RichTextExportRequest();
        var paragraph = request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0];
        paragraph.MarginLeftEmu = 914_400;
        paragraph.IndentEmu = -228_600;
        var authored = Invoke(request);
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var properties = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Paragraph>().First().ParagraphProperties!;
            Assert.Equal(914_400, properties.LeftMargin!.Value);
            Assert.Equal(-228_600, properties.Indent!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var imported = Import(authored.File.ToByteArray());
        Assert.True(imported.Ok, Diagnostics(imported));
        paragraph = imported.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0];
        Assert.Equal(PresentationTextParagraph.LeftMarginOneofCase.MarginLeftEmu, paragraph.LeftMarginCase);
        Assert.Equal(914_400, paragraph.MarginLeftEmu);
        Assert.Equal(PresentationTextParagraph.IndentationOneofCase.IndentEmu, paragraph.IndentationCase);
        Assert.Equal(-228_600, paragraph.IndentEmu);

        paragraph.MarginLeftEmu = 1_143_000;
        paragraph.IndentEmu = -285_750;
        var edited = Export(imported.Artifact);
        Assert.True(edited.Ok, Diagnostics(edited));
        var reimported = Import(edited.File.ToByteArray());
        paragraph = reimported.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0];
        Assert.Equal(1_143_000, paragraph.MarginLeftEmu);
        Assert.Equal(-285_750, paragraph.IndentEmu);

        paragraph.NoMarginLeft = true;
        paragraph.NoIndent = true;
        var deleted = Export(reimported.Artifact);
        Assert.True(deleted.Ok, Diagnostics(deleted));
        using (var stream = new MemoryStream(deleted.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var properties = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Paragraph>().First().ParagraphProperties!;
            Assert.Null(properties.LeftMargin);
            Assert.Null(properties.Indent);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }
    }

    [Fact]
    public void ParagraphSpacingAuthorsImportsEditsAndDeletesWithoutChangingUnits()
    {
        var request = RichTextExportRequest();
        var paragraph = request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0];
        paragraph.LineSpacingMultiplier = 1.25;
        paragraph.SpaceBeforePoints = 12;
        paragraph.SpaceAfterMultiplier = 0.5;
        var authored = Invoke(request);
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var properties = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Paragraph>().First().ParagraphProperties!;
            Assert.Equal(125_000, properties.GetFirstChild<A.LineSpacing>()!.GetFirstChild<A.SpacingPercent>()!.Val!.Value);
            Assert.Equal(1_200, properties.GetFirstChild<A.SpaceBefore>()!.GetFirstChild<A.SpacingPoints>()!.Val!.Value);
            Assert.Equal(50_000, properties.GetFirstChild<A.SpaceAfter>()!.GetFirstChild<A.SpacingPercent>()!.Val!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var imported = Import(authored.File.ToByteArray());
        Assert.True(imported.Ok, Diagnostics(imported));
        paragraph = imported.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0];
        Assert.Equal(PresentationTextParagraph.LineSpacingOneofCase.LineSpacingMultiplier, paragraph.LineSpacingCase);
        Assert.Equal(1.25, paragraph.LineSpacingMultiplier);
        Assert.Equal(PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforePoints, paragraph.SpaceBeforeCase);
        Assert.Equal(12, paragraph.SpaceBeforePoints);
        Assert.Equal(PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterMultiplier, paragraph.SpaceAfterCase);
        Assert.Equal(0.5, paragraph.SpaceAfterMultiplier);

        paragraph.LineSpacingPoints = 18;
        paragraph.SpaceBeforeMultiplier = 0.25;
        paragraph.SpaceAfterPoints = 6;
        var edited = Export(imported.Artifact);
        Assert.True(edited.Ok, Diagnostics(edited));
        var reimported = Import(edited.File.ToByteArray());
        Assert.True(reimported.Ok, Diagnostics(reimported));
        paragraph = reimported.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0];
        Assert.Equal(PresentationTextParagraph.LineSpacingOneofCase.LineSpacingPoints, paragraph.LineSpacingCase);
        Assert.Equal(18, paragraph.LineSpacingPoints);
        Assert.Equal(PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforeMultiplier, paragraph.SpaceBeforeCase);
        Assert.Equal(0.25, paragraph.SpaceBeforeMultiplier);
        Assert.Equal(PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterPoints, paragraph.SpaceAfterCase);
        Assert.Equal(6, paragraph.SpaceAfterPoints);

        paragraph.NoLineSpacing = true;
        paragraph.NoSpaceBefore = true;
        paragraph.NoSpaceAfter = true;
        var deleted = Export(reimported.Artifact);
        Assert.True(deleted.Ok, Diagnostics(deleted));
        using (var stream = new MemoryStream(deleted.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var properties = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Paragraph>().First().ParagraphProperties!;
            Assert.Null(properties.GetFirstChild<A.LineSpacing>());
            Assert.Null(properties.GetFirstChild<A.SpaceBefore>());
            Assert.Null(properties.GetFirstChild<A.SpaceAfter>());
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }
    }

    [Fact]
    public void ParagraphDefaultRunPropertiesAuthorImportEditAndDeleteWhilePreservingUnknownStyle()
    {
        var request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].DefaultRunProperties = new PresentationTextStyle
        {
            Bold = true,
            Italic = false,
            FontSizePoints = 21,
            FontFamily = "Aptos",
            ColorScheme = "accent1",
        };
        var authored = Invoke(request);
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var properties = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Paragraph>().First().ParagraphProperties!;
            var style = properties.GetFirstChild<A.DefaultRunProperties>()!;
            Assert.True(style.Bold!.Value);
            Assert.False(style.Italic!.Value);
            Assert.Equal(2_100, style.FontSize!.Value);
            Assert.Equal("Aptos", style.GetFirstChild<A.LatinFont>()!.Typeface!.Value);
            Assert.Equal(A.SchemeColorValues.Accent1, style.GetFirstChild<A.SolidFill>()!.GetFirstChild<A.SchemeColor>()!.Val!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var source = AddUnmodeledDefaultRunProperties(authored.File.ToByteArray());
        var imported = Import(source);
        Assert.True(imported.Ok, Diagnostics(imported));
        var paragraph = imported.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0];
        Assert.Equal(PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties, paragraph.DefaultRunStyleCase);
        Assert.True(paragraph.DefaultRunProperties.Bold);
        Assert.False(paragraph.DefaultRunProperties.Italic);
        Assert.Equal(21, paragraph.DefaultRunProperties.FontSizePoints);
        Assert.Equal("Aptos", paragraph.DefaultRunProperties.FontFamily);
        Assert.Equal("accent1", paragraph.DefaultRunProperties.ColorScheme);

        paragraph.DefaultRunProperties = new PresentationTextStyle
        {
            Bold = false,
            Italic = true,
            FontSizePoints = 24,
            FontFamily = "Georgia",
            ColorRgb = "2563EB",
        };
        var edited = Export(imported.Artifact);
        Assert.True(edited.Ok, Diagnostics(edited));
        using (var stream = new MemoryStream(edited.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var style = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Paragraph>().First().ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!;
            Assert.False(style.Bold!.Value);
            Assert.True(style.Italic!.Value);
            Assert.Equal(2_400, style.FontSize!.Value);
            Assert.Equal("Georgia", style.GetFirstChild<A.LatinFont>()!.Typeface!.Value);
            Assert.Equal("2563EB", style.GetFirstChild<A.SolidFill>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal(A.TextUnderlineValues.Single, style.Underline!.Value);
            Assert.Equal("Noto Sans CJK SC", style.GetFirstChild<A.EastAsianFont>()!.Typeface!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var reimported = Import(edited.File.ToByteArray());
        paragraph = reimported.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0];
        paragraph.NoDefaultRunProperties = true;
        var deleted = Export(reimported.Artifact);
        Assert.True(deleted.Ok, Diagnostics(deleted));
        using (var stream = new MemoryStream(deleted.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var style = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Paragraph>().First().ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!;
            Assert.Null(style.Bold);
            Assert.Null(style.Italic);
            Assert.Null(style.FontSize);
            Assert.Null(style.GetFirstChild<A.LatinFont>());
            Assert.Null(style.GetFirstChild<A.SolidFill>());
            Assert.Equal(A.TextUnderlineValues.Single, style.Underline!.Value);
            Assert.Equal("Noto Sans CJK SC", style.GetFirstChild<A.EastAsianFont>()!.Typeface!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var transformed = Import(AddTransformedDefaultRunColor(authored.File.ToByteArray()));
        Assert.True(transformed.Ok, Diagnostics(transformed));
        paragraph = transformed.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0];
        Assert.Equal(PresentationTextStyle.ColorOneofCase.None, paragraph.DefaultRunProperties.ColorCase);
        paragraph.DefaultRunProperties.ColorRgb = "FF0000";
        var rejected = Export(transformed.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("unsupported_presentation_edit", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void FieldsBreaksAndTabStopsRoundTripEditAndPreserveResidualProperties()
    {
        var request = RichTextExportRequest();
        var shape = request.Artifact.Presentation.Slides[0].Elements[0].Shape;
        shape.TextBody.Paragraphs.Clear();
        var paragraph = new PresentationTextParagraph();
        paragraph.TabStops.Add(new PresentationTabStop { PositionEmu = 1_143_000, Alignment = "left" });
        paragraph.TabStops.Add(new PresentationTabStop { PositionEmu = 2_476_500, Alignment = "decimal" });
        paragraph.Runs.Add(new PresentationTextRun { Text = "Slide\t" });
        paragraph.Runs.Add(new PresentationTextRun
        {
            Field = new PresentationTextField
            {
                Id = "{11111111-2222-4333-8444-555555555555}",
                Type = "slidenum",
                Text = "1",
            },
            Bold = true,
            ColorRgb = "2563EB",
        });
        paragraph.Runs.Add(new PresentationTextRun { LineBreak = true, FontSizePoints = 13.5 });
        paragraph.Runs.Add(new PresentationTextRun { Text = "Revenue\t42.5" });
        shape.TextBody.Paragraphs.Add(paragraph);
        shape.Text = PptxTextCodec.Flatten(shape.TextBody);

        var authored = Invoke(request);
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var nativeParagraph = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Paragraph>().First();
            Assert.Equal(new[] { "r", "fld", "br", "r" }, nativeParagraph.ChildElements.Where(child => child is A.Run or A.Field or A.Break).Select(child => child.LocalName));
            Assert.Equal("slidenum", nativeParagraph.Elements<A.Field>().Single().Type!.Value);
            Assert.Equal(1350, nativeParagraph.Elements<A.Break>().Single().RunProperties!.FontSize!.Value);
            Assert.Equal(new[] { 1_143_000, 2_476_500 }, nativeParagraph.ParagraphProperties!.GetFirstChild<A.TabStopList>()!.Elements<A.TabStop>().Select(tab => tab.Position!.Value));
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var source = AddUnmodeledInlineProperties(authored.File.ToByteArray());
        var imported = Import(source);
        Assert.True(imported.Ok, Diagnostics(imported));
        var importedShape = imported.Artifact.Presentation.Slides[0].Elements[0].Shape;
        var importedParagraph = importedShape.TextBody.Paragraphs[0];
        Assert.Equal("Slide\t1\nRevenue\t42.5", importedShape.Text);
        Assert.False(importedParagraph.HasNoTabStops);
        Assert.Equal("decimal", importedParagraph.TabStops[1].Alignment);
        Assert.Equal(PresentationTextRun.ContentOneofCase.Text, importedParagraph.Runs[0].ContentCase);
        Assert.Equal(PresentationTextRun.ContentOneofCase.Field, importedParagraph.Runs[1].ContentCase);
        Assert.Equal(PresentationTextRun.ContentOneofCase.LineBreak, importedParagraph.Runs[2].ContentCase);
        Assert.Equal("{11111111-2222-4333-8444-555555555555}", importedParagraph.Runs[1].Field.Id);

        importedParagraph.Runs[1].Field.Text = "2";
        importedParagraph.Runs[2].Italic = true;
        importedParagraph.TabStops[1].PositionEmu = 2_667_000;
        importedShape.Text = PptxTextCodec.Flatten(importedShape.TextBody);
        var preserved = Export(imported.Artifact);
        Assert.True(preserved.Ok, Diagnostics(preserved));
        using (var stream = new MemoryStream(preserved.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var nativeParagraph = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Paragraph>().First();
            var field = nativeParagraph.Elements<A.Field>().Single();
            Assert.Equal("2", field.Text!.Text);
            Assert.Equal(A.TextAlignmentTypeValues.Right, field.ParagraphProperties!.Alignment!.Value);
            var lineBreak = nativeParagraph.Elements<A.Break>().Single();
            Assert.True(lineBreak.RunProperties!.Italic!.Value);
            Assert.Equal(A.TextUnderlineValues.Single, lineBreak.RunProperties.Underline!.Value);
            Assert.Equal(2_667_000, nativeParagraph.ParagraphProperties!.GetFirstChild<A.TabStopList>()!.Elements<A.TabStop>().Last().Position!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var reimported = Import(preserved.File.ToByteArray());
        var reimportedParagraph = reimported.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0];
        Assert.Equal("2", reimportedParagraph.Runs[1].Field.Text);
        Assert.Equal(2_667_000, reimportedParagraph.TabStops[1].PositionEmu);
        reimportedParagraph.TabStops.Clear();
        reimportedParagraph.NoTabStops = true;
        var deletedTabs = Export(reimported.Artifact);
        Assert.True(deletedTabs.Ok, Diagnostics(deletedTabs));
        using (var stream = new MemoryStream(deletedTabs.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
            Assert.Null(package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Paragraph>().First().ParagraphProperties!.GetFirstChild<A.TabStopList>());

        var changedKind = Import(source);
        var changedShape = changedKind.Artifact.Presentation.Slides[0].Elements[0].Shape;
        changedShape.TextBody.Paragraphs[0].Runs[1].Text = "not a field";
        changedShape.Text = PptxTextCodec.Flatten(changedShape.TextBody);
        var rejected = Export(changedKind.Artifact);
        Assert.False(rejected.Ok);
        Assert.Equal("presentation_text_topology_changed", Assert.Single(rejected.Diagnostics).Code);
    }

    [Fact]
    public void InvalidFieldsBreaksAndTabStopsFailClosed()
    {
        var request = RichTextExportRequest();
        var run = request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].Runs[0];
        run.Field = new PresentationTextField { Id = "not-a-guid", Type = "slidenum", Text = "1" };
        var invalidId = Invoke(request);
        Assert.False(invalidId.Ok);
        Assert.Equal("invalid_presentation_text", Assert.Single(invalidId.Diagnostics).Code);

        request = RichTextExportRequest();
        run = request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].Runs[0];
        run.Field = new PresentationTextField { Id = "{11111111-2222-4333-8444-555555555555}", Type = "\u0001", Text = "1" };
        var invalidType = Invoke(request);
        Assert.False(invalidType.Ok);
        Assert.Equal("invalid_presentation_text", Assert.Single(invalidType.Diagnostics).Code);

        request = RichTextExportRequest();
        run = request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].Runs[0];
        run.LineBreak = false;
        var invalidBreak = Invoke(request);
        Assert.False(invalidBreak.Ok);
        Assert.Equal("invalid_presentation_text", Assert.Single(invalidBreak.Diagnostics).Code);

        request = RichTextExportRequest();
        var paragraph = request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0];
        paragraph.NoTabStops = false;
        var invalidDeletion = Invoke(request);
        Assert.False(invalidDeletion.Ok);
        Assert.Equal("invalid_presentation_text", Assert.Single(invalidDeletion.Diagnostics).Code);

        request = RichTextExportRequest();
        paragraph = request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0];
        paragraph.TabStops.Add(new PresentationTabStop { PositionEmu = 100, Alignment = "left" });
        paragraph.NoTabStops = true;
        var contradictoryTabs = Invoke(request);
        Assert.False(contradictoryTabs.Ok);
        Assert.Equal("invalid_presentation_text", Assert.Single(contradictoryTabs.Diagnostics).Code);

        request = RichTextExportRequest();
        paragraph = request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0];
        paragraph.TabStops.Add(new PresentationTabStop { PositionEmu = 200, Alignment = "left" });
        paragraph.TabStops.Add(new PresentationTabStop { PositionEmu = 100, Alignment = "decimal" });
        var unsortedTabs = Invoke(request);
        Assert.False(unsortedTabs.Ok);
        Assert.Equal("invalid_presentation_text", Assert.Single(unsortedTabs.Diagnostics).Code);
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
        Assert.Contains(imported.Artifact.OpaqueOpc.PackageRelationships,
            relationship => relationship.SourcePath == "ppt/slides/slide1.xml" && relationship.Type.EndsWith("/hyperlink", StringComparison.Ordinal));
        Assert.Contains(imported.Artifact.OpaqueOpc.PackageRelationships,
            relationship => relationship.SourcePath == "ppt/slides/slide1.xml" && relationship.Type.EndsWith("/slide", StringComparison.Ordinal));
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
        Assert.Equal("opaque_content_preserved", Assert.Single(preserved.Diagnostics).Code);
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
    public void PictureBulletsAuthorImportAndShareContentAddressedAssetsAcrossSlides()
    {
        var exported = Invoke(PictureBulletExportRequest());
        Assert.True(exported.Ok, Diagnostics(exported));
        using (var stream = new MemoryStream(exported.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var slides = OrderedSlides(package);
            Assert.Equal(2, slides.Length);
            Assert.Equal(slides[0].ImageParts.Single().Uri, slides[1].ImageParts.Single().Uri);
            Assert.Contains(slides[0].ExternalRelationships, relationship =>
                relationship.RelationshipType.EndsWith("/image", StringComparison.Ordinal) &&
                relationship.Uri.OriginalString == "https://example.com/marker.png");
            Assert.Equal(2, slides[0].Slide!.Descendants<A.PictureBullet>().Count());
            Assert.Single(slides[1].Slide!.Descendants<A.PictureBullet>());
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var imported = Import(exported.File.ToByteArray());
        Assert.True(imported.Ok, Diagnostics(imported));
        var asset = Assert.Single(imported.Artifact.Assets);
        Assert.StartsWith("asset/presentation/picture-bullet/", asset.Id);
        Assert.Equal(asset.Id, imported.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].PictureBullet.AssetId);
        Assert.Equal("https://example.com/marker.png", imported.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[1].PictureBullet.Uri);
        Assert.Equal(asset.Id, imported.Artifact.Presentation.Slides[1].Elements[0].Shape.TextBody.Paragraphs[0].PictureBullet.AssetId);
    }

    [Fact]
    public void SourcePreservingPictureBulletEditAddsOnlyTrackedMediaAndKeepsOldGraph()
    {
        var source = Invoke(PictureBulletExportRequest());
        var imported = Import(source.File.ToByteArray());
        Assert.True(imported.Ok, Diagnostics(imported));
        var replacement = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl2nGQAAAAASUVORK5CYII=");
        var replacementId = AddPictureAsset(imported.Artifact, replacement, "image/png");
        var firstShape = imported.Artifact.Presentation.Slides[0].Elements[0].Shape;
        firstShape.TextBody.Paragraphs[0].PictureBullet = new PresentationPictureBullet { AssetId = replacementId };
        firstShape.TextBody.Paragraphs[1].NoBullet = true;
        firstShape.Text = PptxTextCodec.Flatten(firstShape.TextBody);

        var preserved = Export(imported.Artifact);
        Assert.True(preserved.Ok, Diagnostics(preserved));
        using var stream = new MemoryStream(preserved.File.ToByteArray());
        using var package = PresentationDocument.Open(stream, false);
        var slides = OrderedSlides(package);
        Assert.Equal(2, slides[0].ImageParts.Count());
        Assert.Single(slides[1].ImageParts);
        Assert.Contains(slides[0].ExternalRelationships, relationship => relationship.Uri.OriginalString == "https://example.com/marker.png");
        var firstParagraphs = slides[0].Slide!.Descendants<A.Paragraph>().ToArray();
        Assert.NotNull(firstParagraphs[0].ParagraphProperties!.GetFirstChild<A.PictureBullet>());
        Assert.NotNull(firstParagraphs[1].ParagraphProperties!.GetFirstChild<A.NoBullet>());
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
    }

    [Fact]
    public void InvalidPictureBulletAssetsAndUrisFailClosed()
    {
        var request = PictureBulletExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].PictureBullet =
            new PresentationPictureBullet { AssetId = "asset/presentation/picture-bullet/missing" };
        var missing = Invoke(request);
        Assert.False(missing.Ok);
        Assert.Equal("invalid_presentation_asset", Assert.Single(missing.Diagnostics).Code);

        request = PictureBulletExportRequest();
        request.Artifact.Assets[0].Sha256 = new string('0', 64);
        var tampered = Invoke(request);
        Assert.False(tampered.Ok);
        Assert.Equal("invalid_presentation_asset", Assert.Single(tampered.Diagnostics).Code);

        request = PictureBulletExportRequest();
        request.Artifact.Assets.Clear();
        var unsafeSvg = System.Text.Encoding.UTF8.GetBytes("<svg xmlns=\"http://www.w3.org/2000/svg\"><style>@import url(https://example.com/style.css)</style></svg>");
        var unsafeId = AddPictureAsset(request.Artifact, unsafeSvg, "image/svg+xml");
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].PictureBullet =
            new PresentationPictureBullet { AssetId = unsafeId };
        var unsafeResult = Invoke(request);
        Assert.False(unsafeResult.Ok);
        Assert.Equal("invalid_presentation_asset", Assert.Single(unsafeResult.Diagnostics).Code);

        request = PictureBulletExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[1].PictureBullet =
            new PresentationPictureBullet { Uri = "file:///tmp/marker.png" };
        var forbiddenUri = Invoke(request);
        Assert.False(forbiddenUri.Ok);
        Assert.Equal("invalid_presentation_asset", Assert.Single(forbiddenUri.Diagnostics).Code);
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
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].BulletColorScheme = "phClr";
        var invalidSchemeColor = Invoke(request);
        Assert.False(invalidSchemeColor.Ok);
        Assert.Equal("invalid_presentation_color", Assert.Single(invalidSchemeColor.Diagnostics).Code);

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

        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].MarginLeftEmu = 51_206_401;
        var invalidMargin = Invoke(request);
        Assert.False(invalidMargin.Ok);
        Assert.Equal("invalid_presentation_text", Assert.Single(invalidMargin.Diagnostics).Code);

        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].NoIndent = false;
        var invalidIndentDeletion = Invoke(request);
        Assert.False(invalidIndentDeletion.Ok);
        Assert.Equal("invalid_presentation_text", Assert.Single(invalidIndentDeletion.Diagnostics).Code);

        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].SpaceBeforePoints = 1_584.01;
        var invalidSpacing = Invoke(request);
        Assert.False(invalidSpacing.Ok);
        Assert.Equal("invalid_presentation_text", Assert.Single(invalidSpacing.Diagnostics).Code);

        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].NoSpaceAfter = false;
        var invalidSpacingDeletion = Invoke(request);
        Assert.False(invalidSpacingDeletion.Ok);
        Assert.Equal("invalid_presentation_text", Assert.Single(invalidSpacingDeletion.Diagnostics).Code);

        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].DefaultRunProperties = new PresentationTextStyle();
        var emptyDefaultStyle = Invoke(request);
        Assert.False(emptyDefaultStyle.Ok);
        Assert.Equal("invalid_presentation_text", Assert.Single(emptyDefaultStyle.Diagnostics).Code);

        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].DefaultRunProperties = new PresentationTextStyle { FontSizePoints = 769 };
        var invalidDefaultSize = Invoke(request);
        Assert.False(invalidDefaultSize.Ok);
        Assert.Equal("invalid_presentation_text", Assert.Single(invalidDefaultSize.Diagnostics).Code);

        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].DefaultRunProperties = new PresentationTextStyle { ColorScheme = "phClr" };
        var invalidDefaultColor = Invoke(request);
        Assert.False(invalidDefaultColor.Ok);
        Assert.Equal("invalid_presentation_color", Assert.Single(invalidDefaultColor.Diagnostics).Code);

        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.Paragraphs[0].NoDefaultRunProperties = false;
        var invalidDefaultDeletion = Invoke(request);
        Assert.False(invalidDefaultDeletion.Ok);
        Assert.Equal("invalid_presentation_text", Assert.Single(invalidDefaultDeletion.Diagnostics).Code);
    }

    [Fact]
    public void TextBodyListStylesAuthorImportEditDeleteAndPreserveUnknownProperties()
    {
        var request = RichTextExportRequest();
        var body = request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody;
        var first = new PresentationTextParagraph
        {
            Level = 0,
            Alignment = "left",
            BulletCharacter = "•",
            BulletColorScheme = "accent1",
            MarginLeftEmu = 914_400,
            IndentEmu = -228_600,
            SpaceAfterPoints = 6,
            DefaultRunProperties = new PresentationTextStyle { FontSizePoints = 18, ColorScheme = "tx1" },
        };
        first.TabStops.Add(new PresentationTabStop { PositionEmu = 1_828_800, Alignment = "decimal" });
        var picture = new PresentationTextParagraph
        {
            Level = 2,
            PictureBullet = new PresentationPictureBullet(),
            BulletSizeFollowText = true,
            MarginLeftEmu = 1_828_800,
        };
        body.ListStyles.Add(first);
        body.ListStyles.Add(picture);
        var pictureBytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        picture.PictureBullet.AssetId = AddPictureAsset(request.Artifact, pictureBytes, "image/png");

        var authored = Invoke(request);
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var list = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.ListStyle>().Single();
            var level1 = list.GetFirstChild<A.Level1ParagraphProperties>()!;
            Assert.Null(level1.Level);
            Assert.Equal(914_400, level1.LeftMargin!.Value);
            Assert.Equal(-228_600, level1.Indent!.Value);
            Assert.Equal(A.SchemeColorValues.Accent1, level1.GetFirstChild<A.BulletColor>()!.GetFirstChild<A.SchemeColor>()!.Val!.Value);
            Assert.Equal(1_828_800, level1.GetFirstChild<A.TabStopList>()!.GetFirstChild<A.TabStop>()!.Position!.Value);
            Assert.Equal(1_800, level1.GetFirstChild<A.DefaultRunProperties>()!.FontSize!.Value);
            Assert.NotNull(list.GetFirstChild<A.Level3ParagraphProperties>()!.GetFirstChild<A.PictureBullet>());
            Assert.Single(package.PresentationPart.SlideParts.Single().ImageParts);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var source = AddUnmodeledListStyleProperty(authored.File.ToByteArray());
        var imported = Import(source);
        Assert.True(imported.Ok, Diagnostics(imported));
        var shape = imported.Artifact.Presentation.Slides[0].Elements[0].Shape;
        Assert.Equal(new uint[] { 0, 2 }, shape.TextBody.ListStyles.Select(style => style.Level));
        first = shape.TextBody.ListStyles.Single(style => style.Level == 0);
        Assert.Equal("accent1", first.BulletColorScheme);
        Assert.Equal(18, first.DefaultRunProperties.FontSizePoints);
        Assert.Equal("tx1", first.DefaultRunProperties.ColorScheme);
        Assert.Equal(1_828_800, first.TabStops.Single().PositionEmu);
        Assert.Equal(PresentationTextParagraph.BulletOneofCase.PictureBullet, shape.TextBody.ListStyles.Single(style => style.Level == 2).BulletCase);

        first.BulletCharacter = "◆";
        first.BulletColorRgb = "16A34A";
        first.MarginLeftEmu = 1_143_000;
        shape.TextBody.ListStyles.Remove(shape.TextBody.ListStyles.Single(style => style.Level == 2));
        shape.TextBody.ListStyles.Insert(0, new PresentationTextParagraph
        {
            Level = 8,
            AutoNumber = new PresentationAutoNumberBullet { Scheme = "romanLcPeriod", StartAt = 4 },
            BulletFontFollowText = true,
            LineSpacingMultiplier = 1.25,
        });
        var edited = Export(imported.Artifact);
        Assert.True(edited.Ok, Diagnostics(edited));
        using (var stream = new MemoryStream(edited.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var list = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.ListStyle>().Single();
            var level1 = list.GetFirstChild<A.Level1ParagraphProperties>()!;
            Assert.Equal(381_000, level1.RightMargin!.Value);
            Assert.Equal(1_143_000, level1.LeftMargin!.Value);
            Assert.Equal("◆", level1.GetFirstChild<A.CharacterBullet>()!.Char!.Value);
            Assert.Equal("16A34A", level1.GetFirstChild<A.BulletColor>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Null(list.GetFirstChild<A.Level3ParagraphProperties>());
            Assert.Equal("romanLcPeriod", list.GetFirstChild<A.Level9ParagraphProperties>()!.GetFirstChild<A.AutoNumberedBullet>()!.Type!.InnerText);
            Assert.Single(package.PresentationPart.SlideParts.Single().ImageParts);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var reimported = Import(edited.File.ToByteArray());
        Assert.True(reimported.Ok, Diagnostics(reimported));
        shape = reimported.Artifact.Presentation.Slides[0].Elements[0].Shape;
        Assert.Equal(new uint[] { 0, 8 }, shape.TextBody.ListStyles.Select(style => style.Level));
        shape.TextBody.ListStyles.Clear();
        shape.TextBody.NoListStyles = true;
        var deleted = Export(reimported.Artifact);
        Assert.True(deleted.Ok, Diagnostics(deleted));
        using (var stream = new MemoryStream(deleted.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var list = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.ListStyle>().Single();
            var retainedUnknown = list.GetFirstChild<A.Level1ParagraphProperties>()!;
            Assert.Equal(381_000, retainedUnknown.RightMargin!.Value);
            Assert.Null(retainedUnknown.LeftMargin);
            Assert.Null(retainedUnknown.GetFirstChild<A.CharacterBullet>());
            Assert.Null(list.GetFirstChild<A.Level9ParagraphProperties>());
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }
        var deletedImport = Import(deleted.File.ToByteArray());
        Assert.Empty(deletedImport.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.ListStyles);

        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.ListStyles.Add(new PresentationTextParagraph { Alignment = "left" });
        Assert.Equal("invalid_presentation_text", Assert.Single(Invoke(request).Diagnostics).Code);

        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.ListStyles.Add(new PresentationTextParagraph { Level = 0, Alignment = "left" });
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.ListStyles.Add(new PresentationTextParagraph { Level = 0, Alignment = "right" });
        Assert.Equal("invalid_presentation_text", Assert.Single(Invoke(request).Diagnostics).Code);

        request = RichTextExportRequest();
        var invalidRunStyle = new PresentationTextParagraph { Level = 1, Alignment = "left" };
        invalidRunStyle.Runs.Add(new PresentationTextRun { Text = "not allowed" });
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.ListStyles.Add(invalidRunStyle);
        Assert.Equal("invalid_presentation_text", Assert.Single(Invoke(request).Diagnostics).Code);

        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.NoListStyles = false;
        Assert.Equal("invalid_presentation_text", Assert.Single(Invoke(request).Diagnostics).Code);
    }

    [Fact]
    public void TextBodyPropertiesAuthorImportEditDeleteAndPreserveUnknownProperties()
    {
        var request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.BodyProperties = new PresentationTextBodyProperties
        {
            LeftInsetEmu = 76_200,
            TopInsetEmu = 38_100,
            RightInsetEmu = 114_300,
            BottomInsetEmu = 57_150,
            VerticalAnchor = "center",
            Wrap = "none",
            AutoFitMode = "shrinkText",
            RotationAngle60000 = 900_000,
            VerticalTextMode = "vertical",
            VerticalOverflowMode = "ellipsis",
            HorizontalOverflowMode = "clip",
            Columns = 2,
            ColumnSpacingEmu = 171_450,
            RightToLeftColumns = true,
            Upright = true,
        };
        var authored = Invoke(request);
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var properties = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.BodyProperties>().Single();
            Assert.Equal(76_200, properties.LeftInset!.Value);
            Assert.Equal(38_100, properties.TopInset!.Value);
            Assert.Equal(114_300, properties.RightInset!.Value);
            Assert.Equal(57_150, properties.BottomInset!.Value);
            Assert.Equal(A.TextAnchoringTypeValues.Center, properties.Anchor!.Value);
            Assert.Equal(A.TextWrappingValues.None, properties.Wrap!.Value);
            Assert.Equal(900_000, properties.Rotation!.Value);
            Assert.Equal(A.TextVerticalValues.Vertical, properties.Vertical!.Value);
            Assert.Equal(A.TextVerticalOverflowValues.Ellipsis, properties.VerticalOverflow!.Value);
            Assert.Equal(A.TextHorizontalOverflowValues.Clip, properties.HorizontalOverflow!.Value);
            Assert.Equal(2, properties.ColumnCount!.Value);
            Assert.Equal(171_450, properties.ColumnSpacing!.Value);
            Assert.True(properties.RightToLeftColumns!.Value);
            Assert.True(properties.UpRight!.Value);
            Assert.NotNull(properties.GetFirstChild<A.NormalAutoFit>());
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var source = AddUnmodeledBodyProperties(authored.File.ToByteArray());
        var imported = Import(source);
        Assert.True(imported.Ok, Diagnostics(imported));
        var shape = imported.Artifact.Presentation.Slides[0].Elements[0].Shape;
        var bodyProperties = shape.TextBody.BodyProperties;
        Assert.Equal(76_200, bodyProperties.LeftInsetEmu);
        Assert.Equal(38_100, bodyProperties.TopInsetEmu);
        Assert.Equal(114_300, bodyProperties.RightInsetEmu);
        Assert.Equal(57_150, bodyProperties.BottomInsetEmu);
        Assert.Equal("center", bodyProperties.VerticalAnchor);
        Assert.Equal("none", bodyProperties.Wrap);
        Assert.Equal("shrinkText", bodyProperties.AutoFitMode);
        Assert.Equal(900_000, bodyProperties.RotationAngle60000);
        Assert.Equal("vertical", bodyProperties.VerticalTextMode);
        Assert.Equal("ellipsis", bodyProperties.VerticalOverflowMode);
        Assert.Equal("clip", bodyProperties.HorizontalOverflowMode);
        Assert.Equal(2u, bodyProperties.Columns);
        Assert.Equal(171_450, bodyProperties.ColumnSpacingEmu);
        Assert.True(bodyProperties.RightToLeftColumns);
        Assert.True(bodyProperties.Upright);

        bodyProperties.LeftInsetEmu = 152_400;
        bodyProperties.TopInsetEmu = 95_250;
        bodyProperties.NoRightInset = true;
        bodyProperties.BottomInsetEmu = 66_675;
        bodyProperties.VerticalAnchor = "bottom";
        bodyProperties.Wrap = "square";
        bodyProperties.AutoFitMode = "resizeShape";
        bodyProperties.RotationAngle60000 = -1_800_000;
        bodyProperties.VerticalTextMode = "vertical270";
        bodyProperties.VerticalOverflowMode = "clip";
        bodyProperties.HorizontalOverflowMode = "overflow";
        bodyProperties.Columns = 3;
        bodyProperties.ColumnSpacingEmu = 228_600;
        bodyProperties.RightToLeftColumns = false;
        bodyProperties.Upright = false;
        var edited = Export(imported.Artifact);
        Assert.True(edited.Ok, Diagnostics(edited));
        using (var stream = new MemoryStream(edited.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var properties = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.BodyProperties>().Single();
            Assert.Equal(152_400, properties.LeftInset!.Value);
            Assert.Equal(95_250, properties.TopInset!.Value);
            Assert.Null(properties.RightInset);
            Assert.Equal(66_675, properties.BottomInset!.Value);
            Assert.Equal(A.TextAnchoringTypeValues.Bottom, properties.Anchor!.Value);
            Assert.Equal(A.TextWrappingValues.Square, properties.Wrap!.Value);
            Assert.NotNull(properties.GetFirstChild<A.ShapeAutoFit>());
            Assert.Equal(-1_800_000, properties.Rotation!.Value);
            Assert.Equal(A.TextVerticalValues.Vertical270, properties.Vertical!.Value);
            Assert.Equal(A.TextVerticalOverflowValues.Clip, properties.VerticalOverflow!.Value);
            Assert.Equal(A.TextHorizontalOverflowValues.Overflow, properties.HorizontalOverflow!.Value);
            Assert.Equal(3, properties.ColumnCount!.Value);
            Assert.Equal(228_600, properties.ColumnSpacing!.Value);
            Assert.False(properties.RightToLeftColumns!.Value);
            Assert.False(properties.UpRight!.Value);
            Assert.True(properties.AnchorCenter!.Value);
            Assert.True(properties.ForceAntiAlias!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var reimported = Import(edited.File.ToByteArray());
        Assert.True(reimported.Ok, Diagnostics(reimported));
        bodyProperties = reimported.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.BodyProperties;
        bodyProperties.NoLeftInset = true;
        bodyProperties.NoTopInset = true;
        bodyProperties.NoBottomInset = true;
        bodyProperties.NoVerticalAnchor = true;
        bodyProperties.NoWrap = true;
        bodyProperties.NoAutoFitMode = true;
        bodyProperties.NoRotation = true;
        bodyProperties.NoVerticalTextMode = true;
        bodyProperties.NoVerticalOverflowMode = true;
        bodyProperties.NoHorizontalOverflowMode = true;
        bodyProperties.NoColumns = true;
        bodyProperties.NoColumnSpacing = true;
        bodyProperties.NoColumnDirection = true;
        bodyProperties.NoUpright = true;
        var deleted = Export(reimported.Artifact);
        Assert.True(deleted.Ok, Diagnostics(deleted));
        using (var stream = new MemoryStream(deleted.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var properties = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.BodyProperties>().Single();
            Assert.Null(properties.LeftInset);
            Assert.Null(properties.TopInset);
            Assert.Null(properties.RightInset);
            Assert.Null(properties.BottomInset);
            Assert.Null(properties.Anchor);
            Assert.Null(properties.Wrap);
            Assert.DoesNotContain(properties.ChildElements, child => child is A.NoAutoFit or A.NormalAutoFit or A.ShapeAutoFit);
            Assert.Null(properties.Rotation);
            Assert.Null(properties.Vertical);
            Assert.Null(properties.VerticalOverflow);
            Assert.Null(properties.HorizontalOverflow);
            Assert.Null(properties.ColumnCount);
            Assert.Null(properties.ColumnSpacing);
            Assert.Null(properties.RightToLeftColumns);
            Assert.Null(properties.UpRight);
            Assert.True(properties.AnchorCenter!.Value);
            Assert.True(properties.ForceAntiAlias!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var attributedSource = AddAttributedNormalAutoFit(authored.File.ToByteArray());
        var attributedImport = Import(attributedSource);
        Assert.True(attributedImport.Ok, Diagnostics(attributedImport));
        bodyProperties = attributedImport.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.BodyProperties;
        Assert.Equal(PresentationTextBodyProperties.AutoFitOneofCase.None, bodyProperties.AutoFitCase);
        bodyProperties.VerticalAnchor = "top";
        var attributedPreserved = Export(attributedImport.Artifact);
        Assert.True(attributedPreserved.Ok, Diagnostics(attributedPreserved));
        using (var stream = new MemoryStream(attributedPreserved.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var normal = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.NormalAutoFit>().Single();
            Assert.Equal(90_000, normal.FontScale!.Value);
            Assert.Equal(10_000, normal.LineSpaceReduction!.Value);
        }
        bodyProperties.AutoFitMode = "resizeShape";
        var attributedRejected = Export(attributedImport.Artifact);
        Assert.False(attributedRejected.Ok);
        Assert.Equal("unsupported_presentation_edit", Assert.Single(attributedRejected.Diagnostics).Code);

        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.BodyProperties = new PresentationTextBodyProperties { LeftInsetEmu = -1 };
        Assert.Equal("invalid_presentation_text", Assert.Single(Invoke(request).Diagnostics).Code);
        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.BodyProperties = new PresentationTextBodyProperties { NoLeftInset = false };
        Assert.Equal("invalid_presentation_text", Assert.Single(Invoke(request).Diagnostics).Code);
        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.BodyProperties = new PresentationTextBodyProperties { VerticalAnchor = "middle" };
        Assert.Equal("invalid_presentation_text", Assert.Single(Invoke(request).Diagnostics).Code);
        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.BodyProperties = new PresentationTextBodyProperties { AutoFitMode = "stretchText" };
        Assert.Equal("invalid_presentation_text", Assert.Single(Invoke(request).Diagnostics).Code);
        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.BodyProperties = new PresentationTextBodyProperties { RotationAngle60000 = 21_600_001 };
        Assert.Equal("invalid_presentation_text", Assert.Single(Invoke(request).Diagnostics).Code);
        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.BodyProperties = new PresentationTextBodyProperties { VerticalTextMode = "wordArtVertical" };
        Assert.Equal("invalid_presentation_text", Assert.Single(Invoke(request).Diagnostics).Code);
        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.BodyProperties = new PresentationTextBodyProperties { VerticalOverflowMode = "fade" };
        Assert.Equal("invalid_presentation_text", Assert.Single(Invoke(request).Diagnostics).Code);
        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.BodyProperties = new PresentationTextBodyProperties { Columns = 17 };
        Assert.Equal("invalid_presentation_text", Assert.Single(Invoke(request).Diagnostics).Code);
        request = RichTextExportRequest();
        request.Artifact.Presentation.Slides[0].Elements[0].Shape.TextBody.BodyProperties = new PresentationTextBodyProperties { ColumnSpacingEmu = -1 };
        Assert.Equal("invalid_presentation_text", Assert.Single(Invoke(request).Diagnostics).Code);
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

    private static CodecRequest PictureBulletExportRequest()
    {
        var request = ExportRequest();
        var embedded = new PresentationTextParagraph
        {
            PictureBullet = new PresentationPictureBullet(),
        };
        embedded.Runs.Add(new PresentationTextRun { Text = "Embedded marker" });
        var external = new PresentationTextParagraph
        {
            PictureBullet = new PresentationPictureBullet { Uri = "https://example.com/marker.png" },
        };
        external.Runs.Add(new PresentationTextRun { Text = "External marker" });
        var body = new PresentationTextBody();
        body.Paragraphs.Add(embedded);
        body.Paragraphs.Add(external);
        var firstShape = request.Artifact.Presentation.Slides[0].Elements[0].Shape;
        firstShape.TextBody = body;
        firstShape.Text = PptxTextCodec.Flatten(body);

        var secondBody = new PresentationTextBody();
        var secondParagraph = new PresentationTextParagraph { PictureBullet = new PresentationPictureBullet() };
        secondParagraph.Runs.Add(new PresentationTextRun { Text = "Shared marker" });
        secondBody.Paragraphs.Add(secondParagraph);
        var secondSlide = new PresentationSlide { Id = "presentation/slide/2", Name = "Shared asset" };
        var secondElement = request.Artifact.Presentation.Slides[0].Elements[0].Clone();
        secondElement.Id = "presentation/slide/2/title";
        secondElement.Shape.TextBody = secondBody;
        secondElement.Shape.Text = PptxTextCodec.Flatten(secondBody);
        secondSlide.Elements.Add(secondElement);
        request.Artifact.Presentation.Slides.Add(secondSlide);

        var bytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        var assetId = AddPictureAsset(request.Artifact, bytes, "image/png");
        embedded.PictureBullet.AssetId = assetId;
        secondParagraph.PictureBullet.AssetId = assetId;
        return request;
    }

    private static string AddPictureAsset(ArtifactEnvelope envelope, byte[] data, string contentType)
    {
        var digest = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(data)).ToLowerInvariant();
        var id = $"asset/presentation/picture-bullet/{digest}";
        envelope.Assets.Add(new Asset
        {
            Id = id,
            FileName = $"picture-bullet-{digest[..16]}.{(contentType == "image/svg+xml" ? "svg" : "png")}",
            ContentType = contentType,
            Data = ByteString.CopyFrom(data),
            Sha256 = digest,
        });
        return id;
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

    private static byte[] AddUnmodeledDefaultRunProperties(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var presentation = PresentationDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var style = presentation.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Paragraph>().First()
                .ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!;
            style.Underline = A.TextUnderlineValues.Single;
            style.AddChild(new A.EastAsianFont { Typeface = "Noto Sans CJK SC" }, true);
        }
        return stream.ToArray();
    }

    private static byte[] AddTransformedDefaultRunColor(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var presentation = PresentationDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var scheme = presentation.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Paragraph>().First()
                .ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!.GetFirstChild<A.SolidFill>()!.GetFirstChild<A.SchemeColor>()!;
            scheme.Append(new A.Tint { Val = 50_000 });
        }
        return stream.ToArray();
    }

    private static byte[] AddUnmodeledListStyleProperty(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var presentation = PresentationDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var list = presentation.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.ListStyle>().Single();
            list.GetFirstChild<A.Level1ParagraphProperties>()!.RightMargin = 381_000;
        }
        return stream.ToArray();
    }

    private static byte[] AddUnmodeledBodyProperties(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var presentation = PresentationDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var properties = presentation.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.BodyProperties>().Single();
            properties.AnchorCenter = true;
            properties.ForceAntiAlias = true;
        }
        return stream.ToArray();
    }

    private static byte[] AddAttributedNormalAutoFit(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var presentation = PresentationDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var normal = presentation.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.NormalAutoFit>().Single();
            normal.FontScale = 90_000;
            normal.LineSpaceReduction = 10_000;
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
            var transformedScheme = new A.SchemeColor { Val = A.SchemeColorValues.Accent1 };
            transformedScheme.Append(new A.Tint { Val = 50_000 });
            pictureProperties.AddChild(new A.BulletColor(transformedScheme), true);
            pictureProperties.AddChild(new A.BulletFont { Typeface = "Wingdings" }, true);
            // A transformed blip remains deliberately outside the modeled
            // picture-bullet slice and exercises fail-closed preservation.
            pictureProperties.AddChild(new A.PictureBullet(
                new A.Blip(new A.AlphaModulationFixed { Amount = 50_000 }) { Embed = "rIdTextBullet1" }), true);
            var imagePart = slidePart.AddImagePart(ImagePartType.Png, "rIdTextBullet1");
            using (var image = new MemoryStream(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=")))
                imagePart.FeedData(image);
            var properties = paragraph.Elements<A.Run>().First().RunProperties!;
            properties.Underline = A.TextUnderlineValues.Single;
            properties.Append(new A.EastAsianFont { Typeface = "Noto Sans CJK SC" });
        }
        return stream.ToArray();
    }

    private static byte[] AddUnmodeledInlineProperties(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var presentation = PresentationDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var paragraph = presentation.PresentationPart!.SlideParts.First().Slide!.Descendants<A.Paragraph>().First();
            var field = paragraph.Elements<A.Field>().Single();
            field.InsertBefore(new A.ParagraphProperties { Alignment = A.TextAlignmentTypeValues.Right }, field.Text);
            var lineBreak = paragraph.Elements<A.Break>().Single();
            lineBreak.RunProperties!.Underline = A.TextUnderlineValues.Single;
        }
        return stream.ToArray();
    }

    private static byte[] AddTemplatePlaceholders(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var package = PresentationDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var masterPart = package.PresentationPart!.SlideMasterParts.Single();
            masterPart.SlideMaster!.CommonSlideData!.ShapeTree!.Append(TemplatePlaceholder(
                2U,
                "Master Title",
                P.PlaceholderValues.Title,
                0U,
                "Master prompt",
                hasCustomPrompt: true,
                rotation: 60_000));
            var layoutPart = masterPart.SlideLayoutParts.Single();
            layoutPart.SlideLayout!.CommonSlideData!.ShapeTree!.Append(TemplatePlaceholder(
                2U,
                "Layout Body",
                P.PlaceholderValues.Body,
                2U,
                "Layout prompt"));
        }
        return stream.ToArray();
    }

    private static byte[] AddUnsupportedPlaceholderRun(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var package = PresentationDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
        {
            var layout = package.PresentationPart!.SlideMasterParts.Single().SlideLayoutParts.Single().SlideLayout!;
            var textBody = layout.Descendants<P.TextBody>().Single();
            textBody.PrependChild(new A.BodyProperties());
        }
        return stream.ToArray();
    }

    private static P.Shape TemplatePlaceholder(
        uint id,
        string name,
        P.PlaceholderValues type,
        uint index,
        string text,
        bool hasCustomPrompt = false,
        int? rotation = null) => new(
        new P.NonVisualShapeProperties(
            new P.NonVisualDrawingProperties { Id = id, Name = name },
            new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
            new P.ApplicationNonVisualDrawingProperties(
                new P.PlaceholderShape { Type = type, Index = index, HasCustomPrompt = hasCustomPrompt })),
        new P.ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = 762_000L, Y = 571_500L },
                new A.Extents { Cx = 6_858_000L, Cy = 1_143_000L }) { Rotation = rotation },
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
            new A.NoFill()),
        new P.TextBody(
            new A.BodyProperties(),
            new A.ListStyle(),
            new A.Paragraph(
                new A.Run(new A.RunProperties { Language = "en-US" }, new A.Text(text)),
                new A.EndParagraphRunProperties { Language = "en-US" })));

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
