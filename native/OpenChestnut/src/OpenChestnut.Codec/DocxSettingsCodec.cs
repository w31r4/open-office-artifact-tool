using DocumentFormat.OpenXml.Packaging;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Owns the bounded document-wide settings that are part of the public model.
// Header/footer topology remains in DocxHeaderFooterCodec; keeping settings
// here prevents unrelated package concerns from growing together.
internal static class DocxSettingsCodec
{
    private const string WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    internal static void Read(MainDocumentPart mainPart, DocumentArtifact document)
    {
        var settings = mainPart.DocumentSettingsPart?.Settings;
        document.EvenAndOddHeaders = Enabled(settings?.GetFirstChild<W.EvenAndOddHeaders>());
        document.UpdateFields = Enabled(settings?.GetFirstChild<W.UpdateFieldsOnOpen>());
        document.TrackRevisions = Enabled(settings?.GetFirstChild<W.TrackRevisions>());
        if (TryReadProtection(settings, out var protection, out _)) document.DocumentProtection = protection;
    }

    internal static void Author(MainDocumentPart mainPart, DocumentArtifact document)
    {
        var protection = document.DocumentProtection is null ? null : CreateProtection(document.DocumentProtection);
        if (!document.EvenAndOddHeaders && !document.UpdateFields && !document.TrackRevisions && protection is null) return;
        var part = mainPart.AddNewPart<DocumentSettingsPart>();
        part.Settings = new W.Settings();
        if (document.EvenAndOddHeaders) part.Settings.AddChild(new W.EvenAndOddHeaders(), true);
        if (document.TrackRevisions) part.Settings.AddChild(new W.TrackRevisions(), true);
        if (document.UpdateFields) part.Settings.AddChild(new W.UpdateFieldsOnOpen { Val = true }, true);
        if (protection is not null) part.Settings.AddChild(protection, true);
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

        TryReadProtection(mainPart.DocumentSettingsPart?.Settings, out _, out var unsupportedProtection);
        if (unsupportedProtection && requested.DocumentProtection is not null)
            throw new CodecException(
                "unsupported_document_protection_edit",
                "Source-preserving DOCX export cannot replace document protection that contains password verifiers, extensions, or unsupported markup.",
                "word/settings.xml");
        if (requested.DocumentProtection is not null) _ = CreateProtection(requested.DocumentProtection);
    }

    internal static void ApplySource(
        MainDocumentPart mainPart,
        DocumentArtifact requested,
        DocxPartContext context)
    {
        var source = new DocumentArtifact();
        Read(mainPart, source);
        AssertSourceBoundSettings(mainPart, requested);
        var protectionChanged = !EqualProtection(source.DocumentProtection, requested.DocumentProtection);
        if (source.UpdateFields == requested.UpdateFields &&
            source.TrackRevisions == requested.TrackRevisions &&
            !protectionChanged) return;

        var part = mainPart.DocumentSettingsPart ?? mainPart.AddNewPart<DocumentSettingsPart>();
        part.Settings ??= new W.Settings();
        Set(part.Settings, requested.UpdateFields, () => new W.UpdateFieldsOnOpen { Val = true });
        Set(part.Settings, requested.TrackRevisions, () => new W.TrackRevisions());
        if (protectionChanged)
        {
            part.Settings.RemoveAllChildren<W.DocumentProtection>();
            if (requested.DocumentProtection is not null)
                part.Settings.AddChild(CreateProtection(requested.DocumentProtection), true);
        }
        part.Settings.Save();
        context.MarkSettingsMutated(part);
    }

    private static bool TryReadProtection(
        W.Settings? settings,
        out DocumentProtectionSettings? result,
        out bool unsupported)
    {
        result = null;
        unsupported = false;
        var elements = settings?.Elements<W.DocumentProtection>().ToArray() ?? [];
        if (elements.Length == 0) return false;
        if (elements.Length != 1)
        {
            unsupported = true;
            return false;
        }

        var element = elements[0];
        var attributes = element.GetAttributes();
        var expected = new HashSet<string>(StringComparer.Ordinal) { "edit", "enforcement", "formatting" };
        if (element.HasChildren || attributes.Count != expected.Count || attributes.Any(attribute =>
                attribute.NamespaceUri != WordprocessingNamespace || !expected.Remove(attribute.LocalName)) ||
            expected.Count != 0 || element.Edit?.Value is not { } edit ||
            element.Enforcement?.Value is not { } enforcement || element.Formatting?.Value is not { } formatting ||
            !TryMode(edit, out var mode))
        {
            unsupported = true;
            return false;
        }

        result = new DocumentProtectionSettings
        {
            Mode = mode,
            Enforcement = enforcement,
            Formatting = formatting,
        };
        return true;
    }

    private static W.DocumentProtection CreateProtection(DocumentProtectionSettings requested)
    {
        var edit = requested.Mode switch
        {
            DocumentProtectionMode.None => W.DocumentProtectionValues.None,
            DocumentProtectionMode.ReadOnly => W.DocumentProtectionValues.ReadOnly,
            DocumentProtectionMode.Comments => W.DocumentProtectionValues.Comments,
            DocumentProtectionMode.TrackedChanges => W.DocumentProtectionValues.TrackedChanges,
            DocumentProtectionMode.Forms => W.DocumentProtectionValues.Forms,
            _ => throw new CodecException(
                "invalid_document_protection",
                "Document protection mode must be none, readOnly, comments, trackedChanges, or forms.",
                "artifact.document.document_protection.mode"),
        };
        return new W.DocumentProtection
        {
            Edit = edit,
            Enforcement = requested.Enforcement,
            Formatting = requested.Formatting,
        };
    }

    private static bool TryMode(W.DocumentProtectionValues edit, out DocumentProtectionMode mode)
    {
        mode = edit == W.DocumentProtectionValues.None ? DocumentProtectionMode.None
            : edit == W.DocumentProtectionValues.ReadOnly ? DocumentProtectionMode.ReadOnly
            : edit == W.DocumentProtectionValues.Comments ? DocumentProtectionMode.Comments
            : edit == W.DocumentProtectionValues.TrackedChanges ? DocumentProtectionMode.TrackedChanges
            : edit == W.DocumentProtectionValues.Forms ? DocumentProtectionMode.Forms
            : DocumentProtectionMode.Unspecified;
        return mode != DocumentProtectionMode.Unspecified;
    }

    private static bool EqualProtection(DocumentProtectionSettings? left, DocumentProtectionSettings? right) =>
        ReferenceEquals(left, right) || left is not null && right is not null &&
        left.Mode == right.Mode && left.Enforcement == right.Enforcement && left.Formatting == right.Formatting;

    private static void Set<T>(W.Settings settings, bool enabled, Func<T> create)
        where T : W.OnOffType
    {
        settings.RemoveAllChildren<T>();
        if (enabled) settings.AddChild(create(), true);
    }

    private static bool Enabled(W.OnOffType? value) => value is not null && value.Val?.Value != false;
}
