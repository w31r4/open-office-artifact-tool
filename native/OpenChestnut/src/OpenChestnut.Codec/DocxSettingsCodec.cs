using DocumentFormat.OpenXml.Packaging;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Owns the bounded document-wide settings that are part of the public model.
// Header/footer topology remains in DocxHeaderFooterCodec; keeping settings
// here prevents unrelated package concerns from growing together.
internal static class DocxSettingsCodec
{
    internal static void Read(MainDocumentPart mainPart, DocumentArtifact document)
    {
        var settings = mainPart.DocumentSettingsPart?.Settings;
        document.EvenAndOddHeaders = Enabled(settings?.GetFirstChild<W.EvenAndOddHeaders>());
        document.UpdateFields = Enabled(settings?.GetFirstChild<W.UpdateFieldsOnOpen>());
        document.TrackRevisions = Enabled(settings?.GetFirstChild<W.TrackRevisions>());
    }

    internal static void Author(MainDocumentPart mainPart, DocumentArtifact document)
    {
        if (!document.EvenAndOddHeaders && !document.UpdateFields && !document.TrackRevisions) return;
        var part = mainPart.AddNewPart<DocumentSettingsPart>();
        part.Settings = new W.Settings();
        if (document.EvenAndOddHeaders) part.Settings.AddChild(new W.EvenAndOddHeaders(), true);
        if (document.TrackRevisions) part.Settings.AddChild(new W.TrackRevisions(), true);
        if (document.UpdateFields) part.Settings.AddChild(new W.UpdateFieldsOnOpen { Val = true }, true);
        part.Settings.Save();
    }

    internal static void AssertSourceBoundSettings(MainDocumentPart mainPart, DocumentArtifact requested)
    {
        var source = new DocumentArtifact();
        Read(mainPart, source);
        if (source.EvenAndOddHeaders != requested.EvenAndOddHeaders)
            throw new CodecException(
                "unsupported_document_header_footer_edit",
                "Source-preserving DOCX export cannot change even-and-odd header activation because it changes header/footer semantics.",
                "word/settings.xml");
    }

    internal static void ApplySource(
        MainDocumentPart mainPart,
        DocumentArtifact requested,
        DocxPartContext context)
    {
        var source = new DocumentArtifact();
        Read(mainPart, source);
        AssertSourceBoundSettings(mainPart, requested);
        if (source.UpdateFields == requested.UpdateFields &&
            source.TrackRevisions == requested.TrackRevisions) return;

        var part = mainPart.DocumentSettingsPart ?? mainPart.AddNewPart<DocumentSettingsPart>();
        part.Settings ??= new W.Settings();
        Set(part.Settings, requested.UpdateFields, () => new W.UpdateFieldsOnOpen { Val = true });
        Set(part.Settings, requested.TrackRevisions, () => new W.TrackRevisions());
        part.Settings.Save();
        context.MarkSettingsMutated(part);
    }

    private static void Set<T>(W.Settings settings, bool enabled, Func<T> create)
        where T : W.OnOffType
    {
        settings.RemoveAllChildren<T>();
        if (enabled) settings.AddChild(create(), true);
    }

    private static bool Enabled(W.OnOffType? value) => value is not null && value.Val?.Value != false;
}
