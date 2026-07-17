using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenChestnut.Codec;

internal sealed record PptxSpeakerNotesChange(string PartPath, string Sha256);

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

    internal static PptxSpeakerNotesChange? ApplySourceBound(
        SlidePart slidePart,
        PresentationSpeakerNotes? requested,
        int slideIndex)
    {
        var original = Read(slidePart);
        if (original is null)
        {
            if (requested is null || requested.Text.Length == 0) return null;
            throw new CodecException(
                "unsupported_presentation_edit",
                $"Source-preserving PPTX export cannot add speaker notes to slide {slideIndex + 1} because the source slide has no notes part.",
                PartPath(slidePart));
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
        return new PptxSpeakerNotesChange(original.Source.PartPath, HashPart(slidePart.NotesSlidePart!));
    }

    internal static string? BuildSourceFree(
        PresentationPart presentationPart,
        ThemePart themePart,
        IReadOnlyList<SlidePart> slideParts,
        IReadOnlyList<PresentationSlide> slides)
    {
        if (!slides.Any(slide => slide.SpeakerNotes is { Text.Length: > 0 })) return null;

        const string notesMasterRelationshipId = "rIdNotesMaster1";
        var notesMasterPart = presentationPart.AddNewPart<NotesMasterPart>(notesMasterRelationshipId);
        notesMasterPart.AddPart(themePart, "rIdTheme1");
        notesMasterPart.NotesMaster = new P.NotesMaster(
            new P.CommonSlideData(BasicShapeTree()),
            BasicColorMap(),
            new P.HeaderFooter(),
            new P.NotesStyle());

        for (var index = 0; index < slides.Count; index++)
        {
            var notes = slides[index].SpeakerNotes;
            if (notes is null || notes.Text.Length == 0) continue;
            ValidateText(notes.Text);
            var slidePart = slideParts[index];
            var notesPart = slidePart.AddNewPart<NotesSlidePart>("rIdNotes1");
            notesPart.AddPart(notesMasterPart, "rIdNotesMaster1");
            notesPart.AddPart(slidePart, "rIdSlide1");
            var shapeTree = BasicShapeTree();
            shapeTree.Append(NotesBodyShape(notes.Text));
            notesPart.NotesSlide = new P.NotesSlide(
                new P.CommonSlideData(shapeTree),
                new P.ColorMapOverride(new A.MasterColorMapping()));
            notesPart.NotesSlide.Save();
        }
        notesMasterPart.NotesMaster.Save();
        return notesMasterRelationshipId;
    }

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
}
