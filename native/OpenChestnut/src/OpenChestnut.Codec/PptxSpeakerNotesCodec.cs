using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenChestnut.Codec;

internal sealed record PptxSpeakerNotesChange(
    IReadOnlyList<string> ChangedPartPaths,
    IReadOnlyList<string> AddedPartPaths,
    IReadOnlyList<string> AddedRelationshipKeys,
    IReadOnlyDictionary<string, string> ReplacedPartHashes);

// Owns the deliberately narrow speaker-notes contract: plain text in the
// notes body placeholder. Notes masters, layout shapes, styling, and unknown
// relationships stay in the source package and are never projected as if they
// were safely editable.
internal static class PptxSpeakerNotesCodec
{
    private const int MaxTextCharacters = 1_000_000;

    internal static PresentationSpeakerNotes? Read(SlidePart slidePart)
    {
        var notesPart = slidePart.NotesSlidePart;
        if (notesPart?.NotesSlide is not { } notesRoot) return null;
        var body = BodyShape(notesRoot);
        var text = body?.TextBody is { } textBody ? ReadText(textBody) : string.Empty;
        ValidateText(text);
        return new PresentationSpeakerNotes
        {
            Text = text,
            Source = new PresentationSpeakerNotesSourceBinding
            {
                PartPath = PartPath(notesPart),
                RelationshipId = slidePart.GetIdOfPart(notesPart),
                NotesXmlSha256 = HashElement(notesRoot),
                SemanticSha256 = SemanticHash(text),
                Editable = Supports(notesRoot),
            },
        };
    }

    internal static void Validate(PresentationSpeakerNotes? notes)
    {
        if (notes is null) return;
        ValidateText(notes.Text);
    }

    internal static bool CanAddSourceBound(PresentationPart presentationPart, SlidePart slidePart)
    {
        if (slidePart.NotesSlidePart is not null ||
            presentationPart.Presentation is not { } presentation ||
            presentation.SlideMasterIdList is null ||
            presentation.SlideIdList is null ||
            presentation.NotesSize is null ||
            !presentationPart.SlideParts.Contains(slidePart) ||
            !TryResolveNotesMaster(presentationPart, out var notesMasterPart))
            return false;

        return notesMasterPart is not null
            ? ReusableNotesMaster(notesMasterPart)
            : CanonicalThemePart(presentationPart) is not null;
    }

    internal static PptxSpeakerNotesChange? ApplySourceBound(
        PresentationPart presentationPart,
        SlidePart slidePart,
        PresentationSpeakerNotes? requested,
        int slideIndex)
    {
        var original = Read(slidePart);
        if (original is null)
        {
            if (requested is null || requested.Text.Length == 0) return null;
            if (!CanAddSourceBound(presentationPart, slidePart))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Source-preserving PPTX export cannot add speaker notes to slide {slideIndex + 1} because its presentation notes graph is not safely extensible.",
                    PartPath(slidePart));
            ValidateText(requested.Text);
            return AddSourceBound(presentationPart, slidePart, requested.Text);
        }
        if (requested?.Source is not { } binding)
            throw new CodecException(
                "missing_presentation_notes_binding",
                $"Presentation slide {slideIndex + 1} speaker notes are missing their source binding.",
                original.Source.PartPath);
        if (!binding.PartPath.Equals(original.Source.PartPath, StringComparison.OrdinalIgnoreCase) ||
            !binding.RelationshipId.Equals(original.Source.RelationshipId, StringComparison.Ordinal) ||
            !binding.NotesXmlSha256.Equals(original.Source.NotesXmlSha256, StringComparison.OrdinalIgnoreCase) ||
            !binding.SemanticSha256.Equals(original.Source.SemanticSha256, StringComparison.OrdinalIgnoreCase) ||
            binding.Editable != original.Source.Editable)
            throw new CodecException(
                "presentation_notes_binding_mismatch",
                $"Presentation slide {slideIndex + 1} speaker notes do not match their hash-bound source part.",
                original.Source.PartPath);
        ValidateText(requested.Text);
        if (SemanticHash(requested.Text).Equals(original.Source.SemanticSha256, StringComparison.OrdinalIgnoreCase)) return null;
        if (!binding.Editable || slidePart.NotesSlidePart?.NotesSlide is not { } notesRoot || !Supports(notesRoot))
            throw new CodecException(
                "unsupported_presentation_edit",
                $"Presentation slide {slideIndex + 1} speaker notes are preserved but their rich or irregular notes body is not safely editable by this codec slice.",
                original.Source.PartPath);

        ApplyText(BodyShape(notesRoot)!.TextBody!, requested.Text);
        notesRoot.Save();
        return new PptxSpeakerNotesChange(
            [original.Source.PartPath],
            [],
            [],
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                [original.Source.PartPath] = HashPart(slidePart.NotesSlidePart!),
            });
    }

    internal static string? BuildSourceFree(
        PresentationPart presentationPart,
        ThemePart themePart,
        IReadOnlyList<SlidePart> slideParts,
        IReadOnlyList<PresentationSlide> slides)
    {
        if (!slides.Any(slide => slide.SpeakerNotes is { Text.Length: > 0 })) return null;

        const string notesMasterRelationshipId = "rIdNotesMaster1";
        var notesMasterPart = CreateNotesMasterPart(
            presentationPart,
            themePart,
            notesMasterRelationshipId,
            "rIdTheme1");

        for (var index = 0; index < slides.Count; index++)
        {
            var notes = slides[index].SpeakerNotes;
            if (notes is null || notes.Text.Length == 0) continue;
            ValidateText(notes.Text);
            CreateNotesSlidePart(
                slideParts[index],
                notesMasterPart,
                notes.Text,
                "rIdNotes1",
                "rIdNotesMaster1",
                "rIdSlide1");
        }
        return notesMasterRelationshipId;
    }

    internal static void ValidateSourceBoundOutput(
        PresentationPart sourcePresentationPart,
        PresentationPart outputPresentationPart,
        SlidePart sourceSlidePart,
        SlidePart outputSlidePart,
        PresentationSlide requested,
        int slideIndex)
    {
        var source = Read(sourceSlidePart);
        var output = Read(outputSlidePart);
        var expected = requested.SpeakerNotes;
        if (source is null && (expected is null || expected.Text.Length == 0))
        {
            if (output is not null)
                throw Postwrite(slideIndex, "an unchanged slide unexpectedly gained a NotesSlide", PartPath(outputSlidePart));
            return;
        }
        if (expected is null || output is null || !output.Text.Equals(expected.Text, StringComparison.Ordinal))
            throw Postwrite(slideIndex, "speaker-notes text does not match the requested artifact", PartPath(outputSlidePart));

        if (source is not null)
        {
            if (!output.Source.PartPath.Equals(source.Source.PartPath, StringComparison.OrdinalIgnoreCase) ||
                !output.Source.RelationshipId.Equals(source.Source.RelationshipId, StringComparison.Ordinal))
                throw Postwrite(slideIndex, "an existing NotesSlide changed package identity", output.Source.PartPath);
            return;
        }

        ValidateAddedNotesGraph(
            sourcePresentationPart,
            outputPresentationPart,
            outputSlidePart,
            slideIndex);
    }

    private static PptxSpeakerNotesChange AddSourceBound(
        PresentationPart presentationPart,
        SlidePart slidePart,
        string text)
    {
        var changedPartPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var addedPartPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var addedRelationshipKeys = new HashSet<string>(StringComparer.Ordinal);
        if (!TryResolveNotesMaster(presentationPart, out var notesMasterPart))
            throw new CodecException("unsupported_presentation_edit", "The presentation NotesMaster graph is inconsistent.", PartPath(presentationPart));
        notesMasterPart ??= AddSourceBoundNotesMaster(
            presentationPart,
            changedPartPaths,
            addedPartPaths,
            addedRelationshipKeys);

        var slideRelationshipId = NextRelationshipId(slidePart, "rIdNotes");
        var notesPart = CreateNotesSlidePart(
            slidePart,
            notesMasterPart,
            text,
            slideRelationshipId,
            "rIdNotesMaster1",
            "rIdSlide1");
        var notesPath = PartPath(notesPart);
        var notesRelationshipsPath = RelationshipPartPath(notesPart);
        changedPartPaths.Add(RelationshipPartPath(slidePart));
        changedPartPaths.Add(notesPath);
        changedPartPaths.Add(notesRelationshipsPath);
        changedPartPaths.Add("[Content_Types].xml");
        addedPartPaths.Add(notesPath);
        addedPartPaths.Add(notesRelationshipsPath);
        addedRelationshipKeys.Add(RelationshipKey(slidePart, slideRelationshipId));
        addedRelationshipKeys.Add(RelationshipKey(notesPart, "rIdNotesMaster1"));
        addedRelationshipKeys.Add(RelationshipKey(notesPart, "rIdSlide1"));
        return new PptxSpeakerNotesChange(
            changedPartPaths.ToArray(),
            addedPartPaths.ToArray(),
            addedRelationshipKeys.ToArray(),
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase));
    }

    private static NotesMasterPart AddSourceBoundNotesMaster(
        PresentationPart presentationPart,
        ISet<string> changedPartPaths,
        ISet<string> addedPartPaths,
        ISet<string> addedRelationshipKeys)
    {
        var presentation = presentationPart.Presentation ??
            throw new CodecException("missing_presentation_root", "PPTX package has no Presentation root.", PartPath(presentationPart));
        var themePart = CanonicalThemePart(presentationPart) ??
            throw new CodecException("unsupported_presentation_edit", "The presentation has no canonical SlideMaster ThemePart for a NotesMaster.", PartPath(presentationPart));
        var masterRelationshipId = NextRelationshipId(presentationPart, "rIdNotesMaster");
        var notesMasterPart = CreateNotesMasterPart(
            presentationPart,
            themePart,
            masterRelationshipId,
            "rIdTheme1");
        var notesMasterIdList = new P.NotesMasterIdList(
            new P.NotesMasterId { Id = masterRelationshipId });
        presentation.InsertAfter(notesMasterIdList, presentation.SlideMasterIdList!);
        presentation.Save();

        var masterPath = PartPath(notesMasterPart);
        var masterRelationshipsPath = RelationshipPartPath(notesMasterPart);
        changedPartPaths.Add(PartPath(presentationPart));
        changedPartPaths.Add(RelationshipPartPath(presentationPart));
        changedPartPaths.Add(masterPath);
        changedPartPaths.Add(masterRelationshipsPath);
        changedPartPaths.Add("[Content_Types].xml");
        addedPartPaths.Add(masterPath);
        addedPartPaths.Add(masterRelationshipsPath);
        addedRelationshipKeys.Add(RelationshipKey(presentationPart, masterRelationshipId));
        addedRelationshipKeys.Add(RelationshipKey(notesMasterPart, "rIdTheme1"));
        return notesMasterPart;
    }

    private static NotesMasterPart CreateNotesMasterPart(
        PresentationPart presentationPart,
        ThemePart themePart,
        string relationshipId,
        string themeRelationshipId)
    {
        var notesMasterPart = presentationPart.AddNewPart<NotesMasterPart>(relationshipId);
        notesMasterPart.AddPart(themePart, themeRelationshipId);
        notesMasterPart.NotesMaster = CanonicalNotesMaster();
        notesMasterPart.NotesMaster.Save();
        return notesMasterPart;
    }

    private static NotesSlidePart CreateNotesSlidePart(
        SlidePart slidePart,
        NotesMasterPart notesMasterPart,
        string text,
        string slideRelationshipId,
        string masterRelationshipId,
        string slideBackReferenceRelationshipId)
    {
        var notesPart = slidePart.AddNewPart<NotesSlidePart>(slideRelationshipId);
        notesPart.AddPart(notesMasterPart, masterRelationshipId);
        notesPart.AddPart(slidePart, slideBackReferenceRelationshipId);
        var shapeTree = BasicShapeTree();
        shapeTree.Append(NotesBodyShape(text));
        notesPart.NotesSlide = new P.NotesSlide(
            new P.CommonSlideData(shapeTree),
            new P.ColorMapOverride(new A.MasterColorMapping()));
        notesPart.NotesSlide.Save();
        return notesPart;
    }

    private static P.NotesMaster CanonicalNotesMaster() => new(
        new P.CommonSlideData(BasicShapeTree()),
        BasicColorMap(),
        new P.HeaderFooter(),
        new P.NotesStyle());

    private static bool TryResolveNotesMaster(PresentationPart presentationPart, out NotesMasterPart? notesMasterPart)
    {
        notesMasterPart = null;
        var pairs = presentationPart.Parts
            .Where(pair => pair.OpenXmlPart is NotesMasterPart)
            .Take(2)
            .ToArray();
        var ids = presentationPart.Presentation?.NotesMasterIdList?.Elements<P.NotesMasterId>().ToArray() ?? [];
        if (pairs.Length == 0)
            return ids.Length == 0 && presentationPart.Presentation?.NotesMasterIdList is null;
        if (pairs.Length != 1 || ids.Length != 1 || ids[0].Id?.Value != pairs[0].RelationshipId)
            return false;
        notesMasterPart = (NotesMasterPart)pairs[0].OpenXmlPart;
        return true;
    }

    private static ThemePart? CanonicalThemePart(PresentationPart presentationPart)
    {
        foreach (var masterId in presentationPart.Presentation?.SlideMasterIdList?.Elements<P.SlideMasterId>() ?? [])
        {
            var relationshipId = masterId.RelationshipId?.Value ?? string.Empty;
            if (relationshipId.Length > 0 &&
                presentationPart.GetPartById(relationshipId) is SlideMasterPart { ThemePart: { } themePart })
                return themePart;
        }
        return null;
    }

    private static bool ReusableNotesMaster(NotesMasterPart part) =>
        part.NotesMaster is not null && part.ThemePart is not null;

    private static void ValidateAddedNotesGraph(
        PresentationPart sourcePresentationPart,
        PresentationPart outputPresentationPart,
        SlidePart outputSlidePart,
        int slideIndex)
    {
        if (!TryResolveNotesMaster(outputPresentationPart, out var outputMaster) || outputMaster is null)
            throw Postwrite(slideIndex, "the added NotesSlide has no canonical presentation NotesMaster", PartPath(outputPresentationPart));
        var notesPart = outputSlidePart.NotesSlidePart ??
            throw Postwrite(slideIndex, "the requested NotesSlide is absent", PartPath(outputSlidePart));
        var internalParts = notesPart.Parts.ToArray();
        if (internalParts.Length != 2 ||
            internalParts.Count(pair => pair.OpenXmlPart is NotesMasterPart) != 1 ||
            internalParts.Count(pair => pair.OpenXmlPart is SlidePart) != 1 ||
            notesPart.NotesMasterPart?.Uri != outputMaster.Uri ||
            internalParts.Single(pair => pair.OpenXmlPart is SlidePart).OpenXmlPart.Uri != outputSlidePart.Uri ||
            notesPart.ExternalRelationships.Any() ||
            notesPart.HyperlinkRelationships.Any() ||
            notesPart.DataPartReferenceRelationships.Any() ||
            notesPart.NotesSlide is not { } notesRoot ||
            !Supports(notesRoot))
            throw Postwrite(slideIndex, "the added NotesSlide relationship graph is not canonical", PartPath(notesPart));

        if (!TryResolveNotesMaster(sourcePresentationPart, out var sourceMaster))
            throw Postwrite(slideIndex, "the source NotesMaster graph became inconsistent", PartPath(sourcePresentationPart));
        if (sourceMaster is not null)
        {
            if (sourceMaster.Uri != outputMaster.Uri || HashPart(sourceMaster) != HashPart(outputMaster))
                throw Postwrite(slideIndex, "the existing NotesMaster was not reused byte-for-byte", PartPath(outputMaster));
            return;
        }

        var sourceTheme = CanonicalThemePart(sourcePresentationPart);
        var masterParts = outputMaster.Parts.ToArray();
        var common = outputMaster.NotesMaster?.CommonSlideData;
        var shapeTree = common?.ShapeTree;
        if (sourceTheme is null ||
            outputMaster.ThemePart?.Uri != sourceTheme.Uri ||
            masterParts.Length != 1 || masterParts[0].OpenXmlPart is not ThemePart ||
            outputMaster.ExternalRelationships.Any() ||
            outputMaster.HyperlinkRelationships.Any() ||
            outputMaster.DataPartReferenceRelationships.Any() ||
            outputMaster.NotesMaster is not { ColorMap: not null, HeaderFooter: not null, NotesStyle: not null } ||
            shapeTree is null ||
            shapeTree.Elements<P.NonVisualGroupShapeProperties>().Count() != 1 ||
            shapeTree.Elements<P.GroupShapeProperties>().Count() != 1 ||
            shapeTree.ChildElements.Count != 2)
            throw Postwrite(slideIndex, "the new NotesMaster is not the canonical bounded graph", PartPath(outputMaster));
    }

    private static string NextRelationshipId(OpenXmlPartContainer owner, string stem)
    {
        var used = owner.Parts.Select(pair => pair.RelationshipId)
            .Concat(owner.ExternalRelationships.Select(relationship => relationship.Id))
            .Concat(owner.HyperlinkRelationships.Select(relationship => relationship.Id))
            .Concat(owner.DataPartReferenceRelationships.Select(relationship => relationship.Id))
            .ToHashSet(StringComparer.Ordinal);
        for (var index = 1; index <= 1_000_000; index++)
        {
            var candidate = stem + index;
            if (!used.Contains(candidate)) return candidate;
        }
        throw new CodecException("presentation_relationship_budget_exceeded", "PPTX relationship ID allocation exceeded its bounded search.");
    }

    private static CodecException Postwrite(int slideIndex, string message, string path) =>
        new("presentation_postwrite_notes_mismatch", $"PPTX slide {slideIndex + 1} {message}.", path);

    private static P.Shape NotesBodyShape(string text) => new(
        new P.NonVisualShapeProperties(
            new P.NonVisualDrawingProperties { Id = 2U, Name = "Notes Placeholder 1" },
            new P.NonVisualShapeDrawingProperties(),
            new P.ApplicationNonVisualDrawingProperties(
                new P.PlaceholderShape { Type = P.PlaceholderValues.Body, Index = 1U })),
        new P.ShapeProperties(),
        TextBody(text));

    private static P.TextBody TextBody(string text)
    {
        var body = new P.TextBody(new A.BodyProperties(), new A.ListStyle());
        foreach (var line in Lines(text))
            body.Append(new A.Paragraph(new A.Run(new A.Text(line))));
        return body;
    }

    private static void ApplyText(P.TextBody body, string text)
    {
        var template = body.Elements<A.Paragraph>().FirstOrDefault();
        var paragraphProperties = template?.ParagraphProperties?.CloneNode(true);
        var runProperties = template?.Elements<A.Run>().FirstOrDefault()?.RunProperties?.CloneNode(true);
        var endProperties = template?.GetFirstChild<A.EndParagraphRunProperties>()?.CloneNode(true);
        body.RemoveAllChildren<A.Paragraph>();
        foreach (var line in Lines(text))
        {
            var paragraph = new A.Paragraph();
            if (paragraphProperties is not null) paragraph.Append(paragraphProperties.CloneNode(true));
            var run = new A.Run();
            if (runProperties is not null) run.Append(runProperties.CloneNode(true));
            run.Append(new A.Text(line));
            paragraph.Append(run);
            if (endProperties is not null) paragraph.Append(endProperties.CloneNode(true));
            body.Append(paragraph);
        }
    }

    private static string ReadText(P.TextBody body) => string.Join("\n", body.Elements<A.Paragraph>().Select(ParagraphText));

    private static string ParagraphText(A.Paragraph paragraph)
    {
        var result = new StringBuilder();
        foreach (var child in paragraph.ChildElements)
        {
            if (child is A.Run run) result.Append(run.Text?.Text ?? string.Empty);
            else if (child is A.Break) result.Append('\n');
            else if (child is A.Field field) result.Append(field.Text?.Text ?? string.Empty);
        }
        return result.ToString();
    }

    private static bool Supports(P.NotesSlide notes)
    {
        var bodies = notes.CommonSlideData?.ShapeTree?.Elements<P.Shape>()
            .Where(IsBodyPlaceholder)
            .ToArray() ?? [];
        if (bodies.Length != 1 || bodies[0].TextBody is not { } body) return false;
        foreach (var paragraph in body.Elements<A.Paragraph>())
        {
            if (paragraph.ChildElements.Any(child => child is not A.ParagraphProperties and not A.Run and not A.Break and not A.EndParagraphRunProperties)) return false;
            foreach (var run in paragraph.Elements<A.Run>())
                if (run.ChildElements.Any(child => child is not A.RunProperties and not A.Text)) return false;
        }
        var runStyles = body.Descendants<A.RunProperties>().Select(item => item.OuterXml).Distinct(StringComparer.Ordinal).Take(2).Count();
        return runStyles <= 1;
    }

    private static P.Shape? BodyShape(P.NotesSlide notes) => notes.CommonSlideData?.ShapeTree?.Elements<P.Shape>().FirstOrDefault(IsBodyPlaceholder);

    private static bool IsBodyPlaceholder(P.Shape shape) =>
        shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>()?.Type?.Value == P.PlaceholderValues.Body;

    private static string[] Lines(string text) => text.Split('\n', StringSplitOptions.None);

    private static void ValidateText(string text)
    {
        if (text.Length > MaxTextCharacters)
            throw new CodecException("presentation_notes_budget_exceeded", $"Speaker notes exceed the {MaxTextCharacters}-character budget.");
        if (text.Contains('\r'))
            throw new CodecException("invalid_presentation_notes", "Speaker notes must use LF line endings.");
        if (text.Any(character => character < ' ' && character is not '\n' and not '\t'))
            throw new CodecException("invalid_presentation_notes", "Speaker notes contain XML-forbidden control characters.");
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

    private static string SemanticHash(string text) => Hash(Encoding.UTF8.GetBytes(text));
    private static string HashElement(OpenXmlElement element) => Hash(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string HashPart(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static string RelationshipKey(OpenXmlPart source, string relationshipId) => $"{PartPath(source)}\0{relationshipId}";
    private static string RelationshipPartPath(OpenXmlPart part)
    {
        var path = PartPath(part);
        var separator = path.LastIndexOf('/');
        var directory = separator < 0 ? string.Empty : path[..separator];
        var fileName = separator < 0 ? path : path[(separator + 1)..];
        return directory.Length == 0 ? $"_rels/{fileName}.rels" : $"{directory}/_rels/{fileName}.rels";
    }
}
